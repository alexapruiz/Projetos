VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRelTotalPorDocto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Totais por Tipo de Documento"
   ClientHeight    =   4908
   ClientLeft      =   2352
   ClientTop       =   2640
   ClientWidth     =   4908
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4908
   ScaleWidth      =   4908
   Begin VB.Frame fraPrincipal 
      Caption         =   "Informações de  Movimentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4332
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4452
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2172
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
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   1320
         Width           =   3852
      End
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
         TabIndex        =   2
         Top             =   3120
         Width           =   3708
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   312
         Left            =   600
         TabIndex        =   3
         Top             =   3840
         Width           =   1572
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   312
         Left            =   2280
         TabIndex        =   4
         Top             =   3840
         Width           =   1572
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1932
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
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   2052
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
         TabIndex        =   6
         Top             =   2880
         Width           =   1608
      End
   End
   Begin ComctlLib.ProgressBar pgbProcesso 
      Height          =   132
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   4452
      _ExtentX        =   7853
      _ExtentY        =   233
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "frmRelTotalPorDocto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qryMeses            As New rdoQuery
Dim qryDiasMovto        As New rdoQuery
Dim rsMeses             As rdoResultset
Dim rsDiasMovto         As rdoResultset

Private Sub cmbMeses_Click()

    If cmbMeses.ListIndex <> -1 Then
        Call CarregaListDias
        lstDias.Enabled = True
    End If

End Sub

Private Sub cmdImp_Click()
    
    If cmbMeses.ListIndex = -1 Then
        Beep
        MsgBox "Favor selecionar o mês de movimento !", vbExclamation + vbOKOnly, App.Title
        cmbMeses.SetFocus
        Exit Sub
    End If
    
    If lstDias.SelCount = 0 Then
        Beep
        MsgBox "Favor selecionar o(s) dia(s) de movimento(s) !", vbExclamation + vbOKOnly, App.Title
        lstDias.SetFocus
        Exit Sub
    End If
    
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
    
    pgbProcesso.Visible = False
    lstDias.Enabled = False
    Call CarregaCombo
    
End Sub

Private Sub ImprimeRelatorio()
    
Dim qryResult           As rdoQuery
Dim rsResult            As rdoResultset
Dim QryTimeOut          As Variant
Dim i                   As Integer
Dim lEnvelopes          As Long, lMalotes As Long, lConcessionarias As Long, lTitulosUBB As Long
Dim lTitulosOutrosBco   As Long, lLanctoInterno As Long, lTributos As Long
Dim lDepositos          As Long, lArrecadacao As Long
Dim iCountList          As Integer, iTotalList As Integer
Dim DataFormatada       As String

On Error GoTo Err_ImprimeRelatorio

    'Aumenta timeout devido ao processamento da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300

    'Soma quantas datas selecionadas para geração do relatório
    For i = 0 To lstDias.ListCount - 1
        If lstDias.Selected(i) Then iTotalList = iTotalList + 1
    Next
    
    'Parametrização do progress bar
    iCountList = 0
    pgbProcesso.Value = 0
    pgbProcesso.Min = 0
    pgbProcesso.Max = iTotalList
    pgbProcesso.Visible = True
    
    'Soma os Totais por Tipo de documento
    lEnvelopes = 0: lMalotes = 0: lConcessionarias = 0:     lTitulosUBB = 0
    lTitulosOutrosBco = 0: lLanctoInterno = 0: lTributos = 0
    lDepositos = 0: lArrecadacao = 0
        
    Screen.MousePointer = vbHourglass
    
    For i = 0 To lstDias.ListCount - 1
        
        If lstDias.Selected(i) Then
            'Progress Bar
            iCountList = iCountList + 1
            pgbProcesso.Value = iCountList - 0.5

            If Len(DataFormatada) > 0 Then DataFormatada = DataFormatada & " - "
            DataFormatada = DataFormatada & Mid(lstDias.ItemData(i), 7, 2)
            
            Set qryResult = Geral.Banco.CreateQuery("", "{call TotalDoctosPorTipoDocto (?,?)}")
        
            Set rsResult = Nothing
        
            With qryResult
                'Entra com uma das datas selecionadas no list
                .rdoParameters(0).Value = lstDias.ItemData(i)
                'Entra com todas agências de coleta ou uma única  agência
                If cmbAgenf_Agencia.ListIndex <> 0 Then
                    .rdoParameters(1).Value = cmbAgenf_Agencia.ItemData(cmbAgenf_Agencia.ListIndex)
                Else
                    .rdoParameters(1).Value = Null
                End If
        
                Set rsResult = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
            End With

            'Atualiza Progress Bar
            pgbProcesso.Value = iCountList

            If rsResult.RowCount <> 0 Then
                
                Do While Not rsResult.EOF()
                    Select Case rsResult!TipoDocto
                        Case 1                  'Envelope/Malote
                            If rsResult!IdEnv_Mal = "E" Then
                                lEnvelopes = lEnvelopes + rsResult!Total
                            Else
                                lMalotes = lMalotes + rsResult!Total
                            End If
                        Case 2, 3               'Depósitos
                            lDepositos = lDepositos + rsResult!Total
                        Case 15, 16, 17, 18, 35     'Arrecadação (Darm, Darf, Gare e GPS)
                            lArrecadacao = lArrecadacao + rsResult!Total
                        Case 20, 21, 22, 23     'Concessionárias
                            lConcessionarias = lConcessionarias + rsResult!Total
                        Case 28, 29, 30         'Títulos UBB
                            lTitulosUBB = lTitulosUBB + rsResult!Total
                        Case 31, 12                'Títulos Outros Bancos/ Tit. Outros Bancos Convencional
                            lTitulosOutrosBco = lTitulosOutrosBco + rsResult!Total
                        Case 41                 'Lançamento Interno
                            lLanctoInterno = lLanctoInterno + rsResult!Total
                        Case 24, 25, 26         'Tributos
                            lTributos = lTributos + rsResult!Total
                    End Select
                    rsResult.MoveNext
                Loop
            End If
        End If
    Next
    
    With Principal.RptGeral
        .ReportFileName = App.path & "\RelTotalPorTipoDocumento.rpt"
        
        .Formulas(0) = "AgenciaCentral     = '" & Geral.AgenciaCentral & "'"
        .Formulas(1) = "DataMovimento = '" & DataFormatada & "'"
        .Formulas(2) = "TotalMalotes  = '" & Formata(lMalotes, "I") & "'"
        .Formulas(3) = "TotalEnvelopes = '" & Formata(lEnvelopes, "I") & "'"
        .Formulas(4) = "TotalConcessionarias = '" & Formata(lConcessionarias, "I") & "'"
        .Formulas(5) = "TotalTitulosUBB = '" & Formata(lTitulosUBB, "I") & "'"
        .Formulas(6) = "TotalTitulosOutrosBancos = '" & Formata(lTitulosOutrosBco, "I") & "'"
        .Formulas(7) = "TotalLanctosInterno = '" & Formata(lLanctoInterno, "I") & "'"
        .Formulas(8) = "TotalTributos = '" & Formata(lTributos, "I") & "'"
        .Formulas(9) = "TotalDepositos = '" & Formata(lDepositos, "I") & "'"
        .Formulas(10) = "TotalArrecadacao = '" & Formata(lArrecadacao, "I") & "'"
        
        If cmbAgenf_Agencia.ListIndex <> 0 Then
            .Formulas(11) = "AgenciaColeta = '" & Trim(cmbAgenf_Agencia.List(cmbAgenf_Agencia.ListIndex)) & "'"
        Else
            .Formulas(11) = "AgenciaColeta = '" & "Todas'"
        End If
        .Formulas(12) = "MesAnoMovimento  = '( " & cmbMeses.List(cmbMeses.ListIndex) & " )'"

        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Relatório de Totais Aberto Por Tipo de Documento"
        .Action = 1
    End With

    With Principal.RptGeral
        .ReportFileName = Empty
        .Destination = Empty
        .WindowState = Empty
        .WindowTitle = Empty
        For i = 0 To 12: .Formulas(i) = Empty: Next
    End With

Exit_ImprimeRelatorio:
    Screen.MousePointer = vbDefault
    pgbProcesso.Visible = False
    
    'Retorno timeout default
    Geral.Banco.QueryTimeout = QryTimeOut
    'Fecha RecordSet
    qryResult.Close
    If Not (rsResult Is Nothing) Then Set rsResult = Nothing
        
    Exit Sub


Err_ImprimeRelatorio:

    Screen.MousePointer = vbDefault
    Call TratamentoErro("Não foi possível abrir o relatório.", Err, rdoErrors)
    GoTo Exit_ImprimeRelatorio
    
End Sub
Private Sub CarregaCombo()

On Error GoTo Err_CarregaCombo

    Screen.MousePointer = vbHourglass
    
    Set qryMeses = Geral.Banco.CreateQuery("", "{call ObtemDiasMesesMovimento(?)}")
    Set rsMeses = qryMeses.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If rsMeses.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi localizado movimento, favor verificar com suporte!", vbInformation + vbOKOnly, Me.Caption
        Unload Me
        Exit Sub
    End If
    
    cmbMeses.Clear
    'Carrega combo com todos meses/ano de movimento existente
    Do Until rsMeses.EOF()
        cmbMeses.AddItem Mid(CStr(rsMeses(0).Value), 5, 2) & "/" & _
                        Left(CStr(rsMeses(0).Value), 4)
        cmbMeses.ItemData(cmbMeses.NewIndex) = Mid(CStr(rsMeses(0).Value), 5, 2) & _
                                                Left(CStr(rsMeses(0).Value), 4)
        rsMeses.MoveNext
    Loop
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_CarregaCombo:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível obter os meses de movimento, tente novamente", vbInformation + vbOKOnly, App.Title
    Unload Me

End Sub


Private Sub CarregaListDias()

On Error GoTo Err_CarregaListDias

    Screen.MousePointer = vbHourglass
    
    Set qryDiasMovto = Geral.Banco.CreateQuery("", "{call ObtemDiasMesesMovimento(?)}")
    qryDiasMovto.rdoParameters(0) = Format(cmbMeses.ItemData(cmbMeses.ListIndex), "000000")
    Set rsDiasMovto = qryDiasMovto.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If rsDiasMovto.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível obter os dias de movimento, favor tentar novamente !", vbInformation + vbOKOnly, Me.Caption
        Unload Me
        Exit Sub
    End If
    
    lstDias.Clear
    'Carrega combo com todos meses/ano de movimento existente
    Do Until rsDiasMovto.EOF()
        lstDias.AddItem Right(CStr(rsDiasMovto(0).Value), 2) & "/" & _
                        Mid(CStr(rsDiasMovto(0).Value), 5, 2) & "/" & _
                        Left(CStr(rsDiasMovto(0).Value), 4)
        lstDias.ItemData(lstDias.NewIndex) = rsDiasMovto(0).Value

        rsDiasMovto.MoveNext
    Loop
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_CarregaListDias:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível obter os dias de movimento, tente novamente", vbInformation + vbOKOnly, App.Title
    Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set qryDiasMovto = Nothing
    Set rsDiasMovto = Nothing
    Set qryMeses = Nothing
    Set rsMeses = Nothing

End Sub
