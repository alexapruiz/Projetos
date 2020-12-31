VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmAcompRec 
   Caption         =   "Gráfico de  Acompanhamento de Recepção/Captura de Caixa Expresso / Malote Empresa"
   ClientHeight    =   8532
   ClientLeft      =   420
   ClientTop       =   360
   ClientWidth     =   11496
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8532
   ScaleWidth      =   11496
   Begin VB.PictureBox Picture2 
      Height          =   312
      Left            =   5832
      ScaleHeight     =   264
      ScaleWidth      =   2184
      TabIndex        =   34
      Top             =   636
      Width           =   2232
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "AG. Processadora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   36
         TabIndex        =   35
         Top             =   12
         Width           =   2112
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8124
      ScaleHeight     =   276
      ScaleWidth      =   1116
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   636
      Width           =   1140
      Begin VB.Label lblAgProc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   204
         Left            =   36
         TabIndex        =   33
         Top             =   12
         Width           =   1032
      End
   End
   Begin Crystal.CrystalReport RptAcompRec 
      Left            =   10872
      Top             =   8016
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Caption         =   "Geral"
      Height          =   672
      Left            =   5424
      TabIndex        =   26
      Top             =   7032
      Width           =   1428
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   288
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Captura"
      Height          =   672
      Left            =   2760
      TabIndex        =   20
      Top             =   7032
      Width           =   2592
      Begin VB.OptionButton optFiltro 
         Caption         =   "Malotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1524
         TabIndex        =   22
         Top             =   288
         Width           =   1008
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Envelopes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   144
         TabIndex        =   21
         Top             =   264
         Width           =   1368
      End
   End
   Begin VB.Timer TmpAtualizaGrade 
      Interval        =   65535
      Left            =   9936
      Top             =   7944
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   384
      Left            =   5736
      TabIndex        =   18
      Top             =   8028
      Width           =   1512
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   384
      Left            =   4140
      TabIndex        =   17
      Top             =   8028
      Width           =   1512
   End
   Begin VB.Timer TmpDataHora 
      Interval        =   1000
      Left            =   10308
      Top             =   7944
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4416
      ScaleHeight     =   276
      ScaleWidth      =   1116
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   636
      Width           =   1140
      Begin VB.Label lblDataProc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "17/07/2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   204
         Left            =   36
         TabIndex        =   16
         Top             =   12
         Width           =   1032
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recepção"
      Height          =   672
      Left            =   156
      TabIndex        =   4
      Top             =   7032
      Width           =   2544
      Begin VB.OptionButton optFiltro 
         Caption         =   "Envelopes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   144
         TabIndex        =   6
         Top             =   264
         Width           =   1368
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Malotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1524
         TabIndex        =   5
         Top             =   288
         Width           =   948
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   312
      Left            =   2124
      ScaleHeight     =   264
      ScaleWidth      =   2184
      TabIndex        =   0
      Top             =   636
      Width           =   2232
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Data do Movimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   36
         TabIndex        =   1
         Top             =   12
         Width           =   2112
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Legenda"
      Height          =   840
      Left            =   6924
      TabIndex        =   7
      Top             =   6876
      Width           =   4236
      Begin VB.Shape Shape5 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF00FF&
         Height          =   108
         Left            =   2472
         Shape           =   3  'Circle
         Top             =   648
         Width           =   144
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF00FF&
         Height          =   108
         Left            =   2304
         Shape           =   3  'Circle
         Top             =   648
         Width           =   144
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   108
         Left            =   2484
         Shape           =   3  'Circle
         Top             =   420
         Width           =   144
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   108
         Left            =   2316
         Shape           =   3  'Circle
         Top             =   420
         Width           =   144
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         Height          =   84
         Left            =   2328
         TabIndex        =   30
         Top             =   228
         Width           =   252
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000A&
         Height          =   108
         Left            =   276
         TabIndex        =   29
         Top             =   456
         Width           =   60
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000A&
         Height          =   84
         Left            =   276
         TabIndex        =   28
         Top             =   660
         Width           =   60
      End
      Begin VB.Label Label16 
         Caption         =   "Captura Malote"
         Height          =   192
         Left            =   2700
         TabIndex        =   25
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label14 
         Caption         =   "Captura Envelope"
         Height          =   252
         Left            =   2700
         TabIndex        =   24
         Top             =   396
         Width           =   1500
      End
      Begin VB.Label Label13 
         Caption         =   "Captura Todos"
         Height          =   252
         Left            =   2700
         TabIndex        =   23
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Recepção Envelope"
         Height          =   216
         Index           =   0
         Left            =   564
         TabIndex        =   14
         Top             =   396
         Width           =   1512
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000FF&
         Height          =   84
         Left            =   180
         TabIndex        =   13
         Top             =   456
         Width           =   252
      End
      Begin VB.Label Label2 
         Caption         =   "Recepção Todos"
         Height          =   228
         Index           =   1
         Left            =   564
         TabIndex        =   11
         Top             =   192
         Width           =   1308
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         Height          =   84
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   252
      End
      Begin VB.Label Label6 
         Caption         =   "Recepção Malote"
         Height          =   216
         Left            =   564
         TabIndex        =   9
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label9 
         BackColor       =   &H00004000&
         Height          =   84
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   252
      End
   End
   Begin MSChart20Lib.MSChart GrdAcompRec 
      Height          =   6348
      Left            =   240
      OleObjectBlob   =   "FrmEstatisticaHora.frx":0000
      TabIndex        =   31
      Top             =   780
      Width           =   10908
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H0080FFFF&
      Height          =   108
      Left            =   9240
      Shape           =   3  'Circle
      Top             =   7104
      Width           =   108
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17/07/2000"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   108
      TabIndex        =   3
      Top             =   8076
      Width           =   936
   End
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:30:35"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   1104
      TabIndex        =   2
      Top             =   8076
      Width           =   720
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Acompanhamento de Recepção / Captura  de Caixa Expresso / Malote Empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   168
      TabIndex        =   19
      Top             =   276
      Width           =   11172
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   7764
      Left            =   72
      TabIndex        =   12
      Top             =   96
      Width           =   11304
   End
End
Attribute VB_Name = "FrmAcompRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private qryAcompRec As rdoQuery             'Query de pesquisa (Horas & Quantidade de Capas Recepcionadas)
Private qryAcompCap As rdoQuery             'Query de pesquisa (Horas & Quantidade de Capas Capturadas)
Private RsAcompRec  As rdoResultset         'ResultSet para query qryAcompRec
Private RsAcompCap  As rdoResultset         'ResultSet para query qryAcompCap
Private horaInicial As String               'Guarda a hora Inicial de Recepção
Private hrFinalCap  As String               'Guarda a hora Final   de Captura
Private horaFinal   As String               'Guarda a hora Final   de Recepção
Private CountEnv    As Integer              'Contador de Envelopes
Private CountEnvCap As Integer              'Contador de Envelopes - Captura
Private CountMal    As Integer              'Contador de Malotes
Private CountMalCap As Integer              'Contador de Malotes - Captura
Private CountEnvMal As Integer              'Contador de Envelope + Malotes
Private CtEnvMalCap As Integer              'Contador de Envelope + Malotes - Capturados
Private CountHoras  As Integer              'Contador de Horas
Private CountCapas  As Integer              'Contador de Capas
Private CtCapasCap  As Integer              'Contador de Capas - Capturadas
Private TipoGrade   As Integer              'Tipo de Grade escolhida 1 / 2 / 3
Private Contador    As Integer              'Variável auxiliar contadora de Loop

'Type Array - Valores da Base de Dados Recepção
Private Type QtdEnvMalHoras
    Periodo(1 To 48)     As String          'Período formatado 1ª e 2ª parte
    PerForm1(1 To 48)    As String          'Período formatado 1ª parte
    PerForm2(1 To 48)    As String          'Período formatado 2ª parte
    Quantidade(1 To 48)  As Integer         'Quantidade de capas
    TipoEnv(1 To 48)     As String          'Tipo de Capa - Envelope ou Malote
End Type
Private Horas As QtdEnvMalHoras

'Type Array - Valores da Base de Dados Captura
Private Type QtdEnvMalHorasCap
    Periodo(1 To 48)     As String          'Período formatado 1ª e 2ª parte
    PerForm1(1 To 48)    As String          'Período formatado 1ª parte
    PerForm2(1 To 48)    As String          'Período formatado 2ª parte
    Quantidade(1 To 48)  As Integer         'Quantidade de capas
    TipoEnv(1 To 48)     As String          'Tipo de Capa - Envelope ou Malote
End Type
Private HorasCap As QtdEnvMalHorasCap

'Type Array - Valores Formatados para Impressão
Private Type FormataRelhoras
    Periodo(1 To 48)     As String          'Período formatado 1ª parte
    Formata(1 To 48)     As String          'Período formatado 2ª parte
    QtdEnv(1 To 48)      As String          'Quantidade de Envelope Recepção
    QtdMal(1 To 48)      As String          'Quantidade de Malote   Recepçao
    QtdEnvCap(1 To 48)   As String          'Quantidade de Malote   Captura
    QtdMalCap(1 To 48)   As String          'Quantidade de Malote   Captura
    SomaEnv              As Long            'Somatória de Envelopes Recepção
    SomaMal              As Long            'Somatória de Malotes   Recepção
    SomaEnvCap           As Long            'Somatória de Malotes   Captura
    SomaMalCap           As Long            'Somatória de Malotes   Captura
    Porcentagem(1 To 48) As String          'Porcentagem de Cada Meia Hora - Recepção
    PorcentCap(1 To 48)  As String          'Porcentagem de Cada Meia Hora - Captura
End Type
Private FormataRelhoras As FormataRelhoras
Private Sub CmdFechar_Click()
'* Sair *'
    Unload Me
End Sub
Private Sub cmdImprimir_Click()

    Screen.MousePointer = vbHourglass
        PrintForm
        
        If TipoGrade = 1 Or TipoGrade = 2 Then
            Call ImprimeRecepcao
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If TipoGrade = 4 Or TipoGrade = 5 Then
            Call ImprimeCaptura
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        Call ImprimeRecepcao
        Call ImprimeCaptura
            
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Form_Activate()

   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(14)

End Sub
Private Sub Form_Load()

    'Formatação da Data de Processamento
    lblDataProc.Caption = Mid$(Geral.DataProcessamento, 7, 2) & "/" & _
                          Mid$(Geral.DataProcessamento, 5, 2) & "/" & _
                          Mid$(Geral.DataProcessamento, 1, 4)
    
    'Formatação de Agência Processadora
    lblAgProc.Caption = Geral.AgenciaCentral
    
    ' Insere 3 Linhas de Referência
    ' 1 - Todos       (Amarelo)
    ' 2 - Envelope    (Vermelho)
    ' 3 - Malote      (Verde)
    ' 4 - Envelope    (Azul)
    ' 5 - Malote      (Rosa)
    ' 6 - Todos       (Azul Claro)
    
    GrdAcompRec.ColumnCount = 6
    
    'Situação Default - Todos
    optFiltro(3).Value = True
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
'* Limpa variáveis, fecha Recordset e Resultset *'

    Set qryAcompRec = Nothing
    RsAcompRec.Close

    ContaString = 0
    ContaQtdString = 0
    StrVirgula = 0
    StrNumVirgula = 0
    Call StrAposVirgula(0)
    FormataRelhoras.SomaEnv = 0
    FormataRelhoras.SomaMal = 0
    
End Sub
Private Sub optFiltro_Click(Index As Integer)
'* Ao escolher uma opção de OptionButton o programa tratará o tipo de Grade escolhida *'
    '- 1 (Envelope Recepção)
    '- 2 (Malote   Recepção)
    '- 3 (Total    Recepção)
    '- 4 (Envelope Captura)
    '- 5 (Malote   Captura)
    '- 6 (Total    Captura)
        
    If Index = 1 Then
       optFiltro(2).Value = False
       optFiltro(3).Value = False
       optFiltro(4).Value = False
       optFiltro(5).Value = False
       LblTitulo.Caption = "Acompanhamento de Recepção de Caixa Expresso / Malote Empresa"
    End If
    
    If Index = 2 Then
       optFiltro(1).Value = False
       optFiltro(3).Value = False
       optFiltro(4).Value = False
       optFiltro(5).Value = False
       LblTitulo.Caption = "Acompanhamento de Recepção de Caixa Expresso / Malote Empresa"
    End If
    
    If Index = 3 Then
       optFiltro(1).Value = False
       optFiltro(2).Value = False
       optFiltro(4).Value = False
       optFiltro(5).Value = False
       LblTitulo.Caption = "Acompanhamento de Recepção / Captura de Caixa Expresso / Malote Empresa"
    End If
    
    If Index = 4 Then
       optFiltro(1).Value = False
       optFiltro(2).Value = False
       optFiltro(3).Value = False
       optFiltro(5).Value = False
       LblTitulo.Caption = "Acompanhamento de Captura de Caixa Expresso / Malote Empresa"
    End If
    
    If Index = 5 Then
       optFiltro(1).Value = False
       optFiltro(2).Value = False
       optFiltro(3).Value = False
       optFiltro(4).Value = False
       LblTitulo.Caption = "Acompanhamento de Captura de Caixa Expresso / Malote Empresa"
    End If
    
    If optFiltro(1).Value Then
       TipoGrade = 1
    ElseIf optFiltro(2).Value Then
       TipoGrade = 2
    ElseIf optFiltro(3).Value Then
       TipoGrade = 3
    ElseIf optFiltro(4).Value Then
       TipoGrade = 4
    ElseIf optFiltro(5).Value Then
       TipoGrade = 5
    End If
     
    Call Atualiza_AcompRec
    
End Sub
Private Sub TmpAtualizaGrade_Timer()
'* Aciona o timer de atualização *'

    If optFiltro(1).Value Then
       TipoGrade = 1
    ElseIf optFiltro(2).Value Then
       TipoGrade = 2
    ElseIf optFiltro(3).Value Then
       TipoGrade = 3
    ElseIf optFiltro(4).Value Then
       TipoGrade = 4
    ElseIf optFiltro(5).Value Then
       TipoGrade = 5
    End If
     
    Call Atualiza_AcompRec
    
End Sub
Private Sub TmpDataHora_Timer()
  '* Traz a hora Atual *'
  lblData.Caption = Format(Now, "dd/mm/yyyy")
  lblHora.Caption = Format(Now, "hh:mm:ss")
End Sub
Sub Atualiza_AcompRec()
    
    If TrataRecepcao = True Then
    
        Call TrataCaptura
        
        If CInt(Mid(horaFinal, 1, 2)) >= CInt(Mid(hrFinalCap, 1, 2)) Then
            GrdAcompRec.RowCount = CalHoras(horaInicial, horaFinal)
        Else
            GrdAcompRec.RowCount = CalHoras(horaInicial, hrFinalCap)
        End If
        
        Call FormataGradeHoras
        Call PreencheGrade
        Call Calcula_Porcentagem_Rec
        Call Calcula_Porcentagem_Cap
        
    End If

End Sub
Function CalHoras(horaInic, horaFin As String) As Integer
'* Esta funcão calcula a diferença entre a hora inicial de recepção e
'  a hora final de recepção *'

Dim iHor As Integer
Dim fHor As Integer

    iHor = Mid(horaInic, 1, 2)
    fHor = Mid(horaFin, 1, 2)

    If fHor >= 0 And fHor < 6 Then
       fHor = fHor + 24
    End If

    CalHoras = ((fHor - iHor) + 1) * 2

End Function
Sub FormataGradeHoras()
'* Esta rotina preencherá o type formataRelHoras com as horas formatadas para impressão
'  e fará a impressão em tela das horas formatadas para a demonstração do Gráfico '*
   
Dim Resultado   As String
Dim horaCol     As String
    
    horaCol = horaInicial
    
    For CountHoras = 1 To GrdAcompRec.RowCount
    
        GrdAcompRec.Row = CountHoras
        
        If CountHoras > 1 Then
            horaCol = CalcSomaTime(horaCol, "00:30")
        End If
        
        If horaCol = "24:00" Then horaCol = "00:00"
        
        GrdAcompRec.RowLabel = horaCol
        
        If (CountHoras) = 1 Then
            FormataRelhoras.Periodo(CountHoras) = horaCol & " as "
        Else
            FormataRelhoras.Periodo(CountHoras) = CalcSomaTime(horaCol, "00:01") & " as "
        End If
        
        Resultado = CalcSomaTime(horaCol, "00:30")
        
        If Resultado = "24:00" Then
            FormataRelhoras.Formata(CountHoras) = "00:00"
        Else
            FormataRelhoras.Formata(CountHoras) = CalcSomaTime(horaCol, "00:30")
        End If

    Next
    
End Sub
Function CalcSomaTime(TInicio, TFim As String) As String
'* Esta função faz a soma de duas horas *'

Dim iHor, iMin, iSeg As Long
Dim fHor, fMin, fSeg As Long
Dim difHor, difMin, difSeg As Long

    iHor = Val(Mid$(TInicio, 1, 2))
    iMin = Val(Mid$(TInicio, 4, 2))
    iSeg = Val(Mid$(TInicio, 7, 2))
    
    fHor = Val(Mid$(TFim, 1, 2))
    fMin = Val(Mid$(TFim, 4, 2))
    fSeg = Val(Mid$(TFim, 7, 2))
    
    difHor = 0
    difMin = 0
    difSeg = 0

    difSeg = fSeg + iSeg
    
    If difSeg >= 60 Then
        difMin = difMin + 1
        difSeg = difSeg - 60
    End If

    difMin = difMin + fMin + iMin

    If difMin >= 60 Then
        difHor = difHor + 1
        difMin = difMin - 60
    End If

    If iHor >= 24 Then
        iHor = iHor - 24
        If iHor < 25 Then
            difHor = 0
        End If
    End If

    difHor = difHor + fHor + iHor

    CalcSomaTime = StrZero(difHor, 2) + ":" + StrZero(difMin, 2)

End Function
Function StrZero(Num, Size As Long) As String

Dim i As Long

    StrZero = LTrim(Trim(str(Num)))
    
    While Len(StrZero) < Size
        StrZero = "0" + StrZero
    Wend

End Function
Sub PreencheGrade()
'* Esta rotina verifica no array horas(Base de Dados)se exite alguma quantidade de capa
'  para uma determinada hora *'
      
    With GrdAcompRec.Plot.SeriesCollection
        
        .Item(1).Pen.Width = 60
        .Item(1).Pen.Style = VtPenStyleDashed
        
        .Item(2).Pen.Width = 60
        .Item(2).Pen.VtColor.Set 1, 100, 1
        .Item(2).Pen.Style = VtPenStyleDashed
        
        .Item(3).Pen.Width = 70
        .Item(3).Pen.Style = VtPenStyleSolid
        
        .Item(4).Pen.Width = 120
        .Item(4).Pen.Style = VtPenStyleDotted
        
        .Item(5).Pen.Width = 130
        .Item(5).Pen.Style = VtPenStyleDotted
        
        .Item(6).Pen.Width = 110
        .Item(6).Pen.Style = VtPenStyleSolid
         
    End With
        
Select Case TipoGrade

 Case 1: '-- Recepção Envelope --'
        
        GrdAcompRec.Plot.SeriesCollection.Item(1).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(2).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(3).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(4).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(5).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(6).ShowLine = False
            Call OpcaoEnvelope

  Case 2:  '-- Recepção Malote --'
        GrdAcompRec.Plot.SeriesCollection.Item(1).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(2).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(3).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(4).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(5).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(6).ShowLine = False
            Call OpcaoMalote
                                      
  Case 3: '--Envelope / Malote / Todos Recepção + Captura --'
        GrdAcompRec.Plot.SeriesCollection.Item(1).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(2).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(3).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(4).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(5).ShowLine = True
        GrdAcompRec.Plot.SeriesCollection.Item(6).ShowLine = True
            Call OpcaoEnvelope
            Call OpcaoMalote
            Call OpcaoTodos
            Call OpcaoEnvelopeCaptura
            Call OpcaoMaloteCaptura
            Call OpcaoTodosCaptura
    
    Case 4: '-- Captura Envelope --'
        GrdAcompRec.Plot.SeriesCollection.Item(1).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(2).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(3).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(5).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(6).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(4).ShowLine = True
            Call OpcaoEnvelopeCaptura
    
    Case 5: '-- Captura Malote --'
        GrdAcompRec.Plot.SeriesCollection.Item(1).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(2).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(3).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(4).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(6).ShowLine = False
        GrdAcompRec.Plot.SeriesCollection.Item(5).ShowLine = True
            Call OpcaoMaloteCaptura
        
  End Select
    
End Sub
Sub Prepara_Relatorio()
'* Prepara o relátorio de recepção - formulas *'

    With RptAcompRec
    
    '---Formatação de Período
        .Formulas(3) = "Periodo1    = '" & FormataRelhoras.Periodo(1) & FormataRelhoras.Formata(1) & "'"
        .Formulas(4) = "Periodo2    = '" & FormataRelhoras.Periodo(2) & FormataRelhoras.Formata(2) & "'"
        .Formulas(5) = "Periodo3    = '" & FormataRelhoras.Periodo(3) & FormataRelhoras.Formata(3) & "'"
        .Formulas(6) = "Periodo4    = '" & FormataRelhoras.Periodo(4) & FormataRelhoras.Formata(4) & "'"
        .Formulas(7) = "Periodo5    = '" & FormataRelhoras.Periodo(5) & FormataRelhoras.Formata(5) & "'"
        .Formulas(8) = "Periodo6    = '" & FormataRelhoras.Periodo(6) & FormataRelhoras.Formata(6) & "'"
        .Formulas(9) = "Periodo7    = '" & FormataRelhoras.Periodo(7) & FormataRelhoras.Formata(7) & "'"
        .Formulas(10) = "Periodo8   = '" & FormataRelhoras.Periodo(8) & FormataRelhoras.Formata(8) & "'"
        .Formulas(11) = "Periodo9   = '" & FormataRelhoras.Periodo(9) & FormataRelhoras.Formata(9) & "'"
        .Formulas(12) = "Periodo10  = '" & FormataRelhoras.Periodo(10) & FormataRelhoras.Formata(10) & "'"
        .Formulas(13) = "Periodo11  = '" & FormataRelhoras.Periodo(11) & FormataRelhoras.Formata(11) & "'"
        .Formulas(14) = "Periodo12  = '" & FormataRelhoras.Periodo(12) & FormataRelhoras.Formata(12) & "'"
        .Formulas(15) = "Periodo13  = '" & FormataRelhoras.Periodo(13) & FormataRelhoras.Formata(13) & "'"
        .Formulas(16) = "Periodo14  = '" & FormataRelhoras.Periodo(14) & FormataRelhoras.Formata(14) & "'"
        .Formulas(17) = "Periodo15  = '" & FormataRelhoras.Periodo(15) & FormataRelhoras.Formata(15) & "'"
        .Formulas(18) = "Periodo16  = '" & FormataRelhoras.Periodo(16) & FormataRelhoras.Formata(16) & "'"
        .Formulas(19) = "Periodo17  = '" & FormataRelhoras.Periodo(17) & FormataRelhoras.Formata(17) & "'"
        .Formulas(20) = "Periodo18  = '" & FormataRelhoras.Periodo(18) & FormataRelhoras.Formata(18) & "'"
        .Formulas(21) = "Periodo19  = '" & FormataRelhoras.Periodo(19) & FormataRelhoras.Formata(19) & "'"
        .Formulas(22) = "Periodo20  = '" & FormataRelhoras.Periodo(20) & FormataRelhoras.Formata(20) & "'"
        .Formulas(23) = "Periodo21  = '" & FormataRelhoras.Periodo(21) & FormataRelhoras.Formata(21) & "'"
        .Formulas(24) = "Periodo22  = '" & FormataRelhoras.Periodo(22) & FormataRelhoras.Formata(22) & "'"
        .Formulas(25) = "Periodo23  = '" & FormataRelhoras.Periodo(23) & FormataRelhoras.Formata(23) & "'"
        .Formulas(26) = "Periodo24  = '" & FormataRelhoras.Periodo(24) & FormataRelhoras.Formata(24) & "'"
        .Formulas(27) = "Periodo25  = '" & FormataRelhoras.Periodo(25) & FormataRelhoras.Formata(25) & "'"
        .Formulas(28) = "Periodo26  = '" & FormataRelhoras.Periodo(26) & FormataRelhoras.Formata(26) & "'"
        .Formulas(29) = "Periodo27  = '" & FormataRelhoras.Periodo(27) & FormataRelhoras.Formata(27) & "'"
        .Formulas(30) = "Periodo28  = '" & FormataRelhoras.Periodo(28) & FormataRelhoras.Formata(28) & "'"
        .Formulas(31) = "Periodo29  = '" & FormataRelhoras.Periodo(29) & FormataRelhoras.Formata(29) & "'"
        .Formulas(32) = "Periodo30  = '" & FormataRelhoras.Periodo(30) & FormataRelhoras.Formata(30) & "'"
        .Formulas(33) = "Periodo31  = '" & FormataRelhoras.Periodo(31) & FormataRelhoras.Formata(31) & "'"
        .Formulas(96) = "Periodo32  = '" & FormataRelhoras.Periodo(32) & FormataRelhoras.Formata(32) & "'"
        .Formulas(97) = "Periodo33  = '" & FormataRelhoras.Periodo(33) & FormataRelhoras.Formata(33) & "'"
        .Formulas(98) = "Periodo34  = '" & FormataRelhoras.Periodo(34) & FormataRelhoras.Formata(34) & "'"
        .Formulas(99) = "Periodo35  = '" & FormataRelhoras.Periodo(35) & FormataRelhoras.Formata(35) & "'"
        .Formulas(100) = "Periodo36 = '" & FormataRelhoras.Periodo(36) & FormataRelhoras.Formata(36) & "'"
        .Formulas(101) = "Periodo37 = '" & FormataRelhoras.Periodo(37) & FormataRelhoras.Formata(37) & "'"
        .Formulas(102) = "Periodo38 = '" & FormataRelhoras.Periodo(38) & FormataRelhoras.Formata(38) & "'"
        .Formulas(103) = "Periodo39 = '" & FormataRelhoras.Periodo(39) & FormataRelhoras.Formata(39) & "'"
        .Formulas(104) = "Periodo40 = '" & FormataRelhoras.Periodo(40) & FormataRelhoras.Formata(40) & "'"
        .Formulas(105) = "Periodo41 = '" & FormataRelhoras.Periodo(41) & FormataRelhoras.Formata(41) & "'"
        .Formulas(106) = "Periodo42 = '" & FormataRelhoras.Periodo(42) & FormataRelhoras.Formata(42) & "'"
        .Formulas(107) = "Periodo43 = '" & FormataRelhoras.Periodo(43) & FormataRelhoras.Formata(43) & "'"
        .Formulas(108) = "Periodo44 = '" & FormataRelhoras.Periodo(44) & FormataRelhoras.Formata(44) & "'"
        .Formulas(109) = "Periodo45 = '" & FormataRelhoras.Periodo(45) & FormataRelhoras.Formata(45) & "'"
        .Formulas(110) = "Periodo46 = '" & FormataRelhoras.Periodo(46) & FormataRelhoras.Formata(46) & "'"
        .Formulas(111) = "Periodo47 = '" & FormataRelhoras.Periodo(47) & FormataRelhoras.Formata(47) & "'"
        .Formulas(112) = "Periodo48 = '" & FormataRelhoras.Periodo(48) & FormataRelhoras.Formata(48) & "'"
     
     '---Formatação de Quantide de Envelopes
        .Formulas(34) = "Envelope1  = '" & FormataRelhoras.QtdEnv(1) & "'"
        .Formulas(35) = "Envelope2  = '" & FormataRelhoras.QtdEnv(2) & "'"
        .Formulas(36) = "Envelope3  = '" & FormataRelhoras.QtdEnv(3) & "'"
        .Formulas(37) = "Envelope4  = '" & FormataRelhoras.QtdEnv(4) & "'"
        .Formulas(38) = "Envelope5  = '" & FormataRelhoras.QtdEnv(5) & "'"
        .Formulas(39) = "Envelope6  = '" & FormataRelhoras.QtdEnv(6) & "'"
        .Formulas(40) = "Envelope7  = '" & FormataRelhoras.QtdEnv(7) & "'"
        .Formulas(41) = "Envelope8  = '" & FormataRelhoras.QtdEnv(8) & "'"
        .Formulas(42) = "Envelope9  = '" & FormataRelhoras.QtdEnv(9) & "'"
        .Formulas(43) = "Envelope10 = '" & FormataRelhoras.QtdEnv(10) & "'"
        .Formulas(44) = "Envelope11 = '" & FormataRelhoras.QtdEnv(11) & "'"
        .Formulas(45) = "Envelope12 = '" & FormataRelhoras.QtdEnv(12) & "'"
        .Formulas(46) = "Envelope13 = '" & FormataRelhoras.QtdEnv(13) & "'"
        .Formulas(47) = "Envelope14 = '" & FormataRelhoras.QtdEnv(14) & "'"
        .Formulas(48) = "Envelope15 = '" & FormataRelhoras.QtdEnv(15) & "'"
        .Formulas(49) = "Envelope16 = '" & FormataRelhoras.QtdEnv(16) & "'"
        .Formulas(50) = "Envelope17 = '" & FormataRelhoras.QtdEnv(17) & "'"
        .Formulas(51) = "Envelope18 = '" & FormataRelhoras.QtdEnv(18) & "'"
        .Formulas(52) = "Envelope19 = '" & FormataRelhoras.QtdEnv(19) & "'"
        .Formulas(53) = "Envelope20 = '" & FormataRelhoras.QtdEnv(20) & "'"
        .Formulas(54) = "Envelope21 = '" & FormataRelhoras.QtdEnv(21) & "'"
        .Formulas(55) = "Envelope22 = '" & FormataRelhoras.QtdEnv(22) & "'"
        .Formulas(56) = "Envelope23 = '" & FormataRelhoras.QtdEnv(23) & "'"
        .Formulas(57) = "Envelope24 = '" & FormataRelhoras.QtdEnv(24) & "'"
        .Formulas(58) = "Envelope25 = '" & FormataRelhoras.QtdEnv(25) & "'"
        .Formulas(59) = "Envelope26 = '" & FormataRelhoras.QtdEnv(26) & "'"
        .Formulas(60) = "Envelope27 = '" & FormataRelhoras.QtdEnv(27) & "'"
        .Formulas(61) = "Envelope28 = '" & FormataRelhoras.QtdEnv(28) & "'"
        .Formulas(62) = "Envelope29 = '" & FormataRelhoras.QtdEnv(29) & "'"
        .Formulas(63) = "Envelope30 = '" & FormataRelhoras.QtdEnv(30) & "'"
        .Formulas(113) = "Envelope31 = '" & FormataRelhoras.QtdEnv(31) & "'"
        .Formulas(114) = "Envelope32 = '" & FormataRelhoras.QtdEnv(32) & "'"
        .Formulas(115) = "Envelope33 = '" & FormataRelhoras.QtdEnv(33) & "'"
        .Formulas(116) = "Envelope34 = '" & FormataRelhoras.QtdEnv(34) & "'"
        .Formulas(117) = "Envelope35 = '" & FormataRelhoras.QtdEnv(35) & "'"
        .Formulas(118) = "Envelope36 = '" & FormataRelhoras.QtdEnv(36) & "'"
        .Formulas(119) = "Envelope37 = '" & FormataRelhoras.QtdEnv(37) & "'"
        .Formulas(120) = "Envelope38 = '" & FormataRelhoras.QtdEnv(38) & "'"
        .Formulas(121) = "Envelope39 = '" & FormataRelhoras.QtdEnv(39) & "'"
        .Formulas(122) = "Envelope40 = '" & FormataRelhoras.QtdEnv(40) & "'"
        .Formulas(123) = "Envelope41 = '" & FormataRelhoras.QtdEnv(41) & "'"
        .Formulas(124) = "Envelope42 = '" & FormataRelhoras.QtdEnv(42) & "'"
        .Formulas(125) = "Envelope43 = '" & FormataRelhoras.QtdEnv(43) & "'"
        .Formulas(126) = "Envelope44 = '" & FormataRelhoras.QtdEnv(44) & "'"
        .Formulas(127) = "Envelope45 = '" & FormataRelhoras.QtdEnv(45) & "'"
        .Formulas(128) = "Envelope46 = '" & FormataRelhoras.QtdEnv(46) & "'"
        .Formulas(129) = "Envelope47 = '" & FormataRelhoras.QtdEnv(47) & "'"
        .Formulas(130) = "Envelope48 = '" & FormataRelhoras.QtdEnv(48) & "'"
    
    '---Formatação de Quantidades de Malotes
        .Formulas(65) = "Malote1   = '" & FormataRelhoras.QtdMal(1) & "'"
        .Formulas(66) = "Malote2   = '" & FormataRelhoras.QtdMal(2) & "'"
        .Formulas(67) = "Malote3   = '" & FormataRelhoras.QtdMal(3) & "'"
        .Formulas(68) = "Malote4   = '" & FormataRelhoras.QtdMal(4) & "'"
        .Formulas(69) = "Malote5   = '" & FormataRelhoras.QtdMal(5) & "'"
        .Formulas(70) = "Malote6   = '" & FormataRelhoras.QtdMal(6) & "'"
        .Formulas(71) = "Malote7   = '" & FormataRelhoras.QtdMal(7) & "'"
        .Formulas(72) = "Malote8   = '" & FormataRelhoras.QtdMal(8) & "'"
        .Formulas(73) = "Malote9   = '" & FormataRelhoras.QtdMal(9) & "'"
        .Formulas(74) = "Malote10  = '" & FormataRelhoras.QtdMal(10) & "'"
        .Formulas(75) = "Malote11  = '" & FormataRelhoras.QtdMal(11) & "'"
        .Formulas(76) = "Malote12  = '" & FormataRelhoras.QtdMal(12) & "'"
        .Formulas(77) = "Malote13  = '" & FormataRelhoras.QtdMal(13) & "'"
        .Formulas(78) = "Malote14  = '" & FormataRelhoras.QtdMal(14) & "'"
        .Formulas(79) = "Malote15  = '" & FormataRelhoras.QtdMal(15) & "'"
        .Formulas(80) = "Malote16  = '" & FormataRelhoras.QtdMal(16) & "'"
        .Formulas(81) = "Malote17  = '" & FormataRelhoras.QtdMal(17) & "'"
        .Formulas(82) = "Malote18  = '" & FormataRelhoras.QtdMal(18) & "'"
        .Formulas(83) = "Malote19  = '" & FormataRelhoras.QtdMal(19) & "'"
        .Formulas(84) = "Malote20  = '" & FormataRelhoras.QtdMal(20) & "'"
        .Formulas(85) = "Malote21  = '" & FormataRelhoras.QtdMal(21) & "'"
        .Formulas(86) = "Malote22  = '" & FormataRelhoras.QtdMal(22) & "'"
        .Formulas(87) = "Malote23  = '" & FormataRelhoras.QtdMal(23) & "'"
        .Formulas(88) = "Malote24  = '" & FormataRelhoras.QtdMal(24) & "'"
        .Formulas(89) = "Malote25  = '" & FormataRelhoras.QtdMal(25) & "'"
        .Formulas(90) = "Malote26  = '" & FormataRelhoras.QtdMal(26) & "'"
        .Formulas(91) = "Malote27  = '" & FormataRelhoras.QtdMal(27) & "'"
        .Formulas(92) = "Malote28  = '" & FormataRelhoras.QtdMal(28) & "'"
        .Formulas(93) = "Malote29  = '" & FormataRelhoras.QtdMal(29) & "'"
        .Formulas(94) = "Malote30  = '" & FormataRelhoras.QtdMal(30) & "'"
        .Formulas(131) = "Malote31 = '" & FormataRelhoras.QtdMal(31) & "'"
        .Formulas(132) = "Malote32 = '" & FormataRelhoras.QtdMal(32) & "'"
        .Formulas(133) = "Malote33 = '" & FormataRelhoras.QtdMal(33) & "'"
        .Formulas(134) = "Malote34 = '" & FormataRelhoras.QtdMal(34) & "'"
        .Formulas(135) = "Malote35 = '" & FormataRelhoras.QtdMal(35) & "'"
        .Formulas(136) = "Malote36 = '" & FormataRelhoras.QtdMal(36) & "'"
        .Formulas(137) = "Malote37 = '" & FormataRelhoras.QtdMal(37) & "'"
        .Formulas(138) = "Malote38 = '" & FormataRelhoras.QtdMal(38) & "'"
        .Formulas(139) = "Malote39 = '" & FormataRelhoras.QtdMal(39) & "'"
        .Formulas(140) = "Malote40 = '" & FormataRelhoras.QtdMal(40) & "'"
        .Formulas(141) = "Malote41 = '" & FormataRelhoras.QtdMal(41) & "'"
        .Formulas(142) = "Malote42 = '" & FormataRelhoras.QtdMal(42) & "'"
        .Formulas(143) = "Malote43 = '" & FormataRelhoras.QtdMal(43) & "'"
        .Formulas(144) = "Malote44 = '" & FormataRelhoras.QtdMal(44) & "'"
        .Formulas(145) = "Malote45 = '" & FormataRelhoras.QtdMal(45) & "'"
        .Formulas(146) = "Malote46 = '" & FormataRelhoras.QtdMal(46) & "'"
        .Formulas(147) = "Malote47 = '" & FormataRelhoras.QtdMal(47) & "'"
        .Formulas(148) = "Malote48 = '" & FormataRelhoras.QtdMal(48) & "'"
      
      '--Total de Capas
        .Formulas(149) = "TotalEnvelope = '" & FormataRelhoras.SomaEnv & "'"
        .Formulas(150) = "TotalMalote= '" & FormataRelhoras.SomaMal & "'"
      
      '--Porcentagem
        .Formulas(151) = "Porcentagem1 = '" & FormataRelhoras.Porcentagem(1) & "'"
        .Formulas(152) = "Porcentagem2 = '" & FormataRelhoras.Porcentagem(2) & "'"
        .Formulas(153) = "Porcentagem3 = '" & FormataRelhoras.Porcentagem(3) & "'"
        .Formulas(154) = "Porcentagem4 = '" & FormataRelhoras.Porcentagem(4) & "'"
        .Formulas(155) = "Porcentagem5 = '" & FormataRelhoras.Porcentagem(5) & "'"
        .Formulas(156) = "Porcentagem6 = '" & FormataRelhoras.Porcentagem(6) & "'"
        .Formulas(157) = "Porcentagem7 = '" & FormataRelhoras.Porcentagem(7) & "'"
        .Formulas(158) = "Porcentagem8 = '" & FormataRelhoras.Porcentagem(8) & "'"
        .Formulas(159) = "Porcentagem9 = '" & FormataRelhoras.Porcentagem(9) & "'"
        .Formulas(160) = "Porcentagem10= '" & FormataRelhoras.Porcentagem(10) & "'"
        .Formulas(161) = "Porcentagem11= '" & FormataRelhoras.Porcentagem(11) & "'"
        .Formulas(162) = "Porcentagem12= '" & FormataRelhoras.Porcentagem(12) & "'"
        .Formulas(163) = "Porcentagem13= '" & FormataRelhoras.Porcentagem(13) & "'"
        .Formulas(164) = "Porcentagem14= '" & FormataRelhoras.Porcentagem(14) & "'"
        .Formulas(165) = "Porcentagem15= '" & FormataRelhoras.Porcentagem(15) & "'"
        .Formulas(166) = "Porcentagem16= '" & FormataRelhoras.Porcentagem(16) & "'"
        .Formulas(167) = "Porcentagem17= '" & FormataRelhoras.Porcentagem(17) & "'"
        .Formulas(168) = "Porcentagem18= '" & FormataRelhoras.Porcentagem(18) & "'"
        .Formulas(169) = "Porcentagem19= '" & FormataRelhoras.Porcentagem(19) & "'"
        .Formulas(171) = "Porcentagem20= '" & FormataRelhoras.Porcentagem(20) & "'"
        .Formulas(172) = "Porcentagem21= '" & FormataRelhoras.Porcentagem(21) & "'"
        .Formulas(173) = "Porcentagem22= '" & FormataRelhoras.Porcentagem(22) & "'"
        .Formulas(174) = "Porcentagem23= '" & FormataRelhoras.Porcentagem(23) & "'"
        .Formulas(175) = "Porcentagem24= '" & FormataRelhoras.Porcentagem(24) & "'"
        .Formulas(176) = "Porcentagem25= '" & FormataRelhoras.Porcentagem(25) & "'"
        .Formulas(177) = "Porcentagem26= '" & FormataRelhoras.Porcentagem(26) & "'"
        .Formulas(178) = "Porcentagem27= '" & FormataRelhoras.Porcentagem(27) & "'"
        .Formulas(179) = "Porcentagem28= '" & FormataRelhoras.Porcentagem(28) & "'"
        .Formulas(181) = "Porcentagem29= '" & FormataRelhoras.Porcentagem(29) & "'"
        .Formulas(182) = "Porcentagem30= '" & FormataRelhoras.Porcentagem(30) & "'"
        .Formulas(183) = "Porcentagem31= '" & FormataRelhoras.Porcentagem(31) & "'"
        .Formulas(184) = "Porcentagem32= '" & FormataRelhoras.Porcentagem(32) & "'"
        .Formulas(185) = "Porcentagem33= '" & FormataRelhoras.Porcentagem(33) & "'"
        .Formulas(186) = "Porcentagem34= '" & FormataRelhoras.Porcentagem(34) & "'"
        .Formulas(187) = "Porcentagem35= '" & FormataRelhoras.Porcentagem(35) & "'"
        .Formulas(188) = "Porcentagem36= '" & FormataRelhoras.Porcentagem(36) & "'"
        .Formulas(189) = "Porcentagem37= '" & FormataRelhoras.Porcentagem(37) & "'"
        .Formulas(191) = "Porcentagem38= '" & FormataRelhoras.Porcentagem(38) & "'"
        .Formulas(192) = "Porcentagem39= '" & FormataRelhoras.Porcentagem(39) & "'"
        .Formulas(193) = "Porcentagem40= '" & FormataRelhoras.Porcentagem(40) & "'"
        .Formulas(194) = "Porcentagem41= '" & FormataRelhoras.Porcentagem(41) & "'"
        .Formulas(195) = "Porcentagem42= '" & FormataRelhoras.Porcentagem(42) & "'"
        .Formulas(196) = "Porcentagem43= '" & FormataRelhoras.Porcentagem(43) & "'"
        .Formulas(197) = "Porcentagem44= '" & FormataRelhoras.Porcentagem(44) & "'"
        .Formulas(198) = "Porcentagem45= '" & FormataRelhoras.Porcentagem(45) & "'"
        .Formulas(199) = "Porcentagem46= '" & FormataRelhoras.Porcentagem(46) & "'"
        .Formulas(200) = "Porcentagem47= '" & FormataRelhoras.Porcentagem(47) & "'"
        .Formulas(201) = "Porcentagem48= '" & FormataRelhoras.Porcentagem(48) & "'"
        .Formulas(202) = "AgProcessadora= '" & Geral.AgenciaCentral & "'"
        
    End With

End Sub
Sub Limpa_Formula()
'Limpa todas as formulas que foram enviadas para o relatório

    With RptAcompRec
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = ""
        .Formulas(5) = ""
        .Formulas(6) = ""
        .Formulas(7) = ""
        .Formulas(8) = ""
        .Formulas(9) = ""
        .Formulas(10) = ""
        .Formulas(11) = ""
        .Formulas(12) = ""
        .Formulas(13) = ""
        .Formulas(14) = ""
        .Formulas(15) = ""
        .Formulas(16) = ""
        .Formulas(17) = ""
        .Formulas(18) = ""
        .Formulas(19) = ""
        .Formulas(20) = ""
        .Formulas(21) = ""
        .Formulas(22) = ""
        .Formulas(23) = ""
        .Formulas(24) = ""
        .Formulas(25) = ""
        .Formulas(26) = ""
        .Formulas(27) = ""
        .Formulas(28) = ""
        .Formulas(29) = ""
        .Formulas(30) = ""
        .Formulas(31) = ""
        .Formulas(32) = ""
        .Formulas(33) = ""
        .Formulas(34) = ""
        .Formulas(35) = ""
        .Formulas(36) = ""
        .Formulas(37) = ""
        .Formulas(38) = ""
        .Formulas(39) = ""
        .Formulas(40) = ""
        .Formulas(41) = ""
        .Formulas(42) = ""
        .Formulas(43) = ""
        .Formulas(44) = ""
        .Formulas(45) = ""
        .Formulas(46) = ""
        .Formulas(47) = ""
        .Formulas(48) = ""
        .Formulas(49) = ""
        .Formulas(50) = ""
        .Formulas(51) = ""
        .Formulas(52) = ""
        .Formulas(53) = ""
        .Formulas(54) = ""
        .Formulas(55) = ""
        .Formulas(56) = ""
        .Formulas(57) = ""
        .Formulas(58) = ""
        .Formulas(59) = ""
        .Formulas(60) = ""
        .Formulas(61) = ""
        .Formulas(62) = ""
        .Formulas(63) = ""
        .Formulas(64) = ""
        .Formulas(65) = ""
        .Formulas(66) = ""
        .Formulas(67) = ""
        .Formulas(68) = ""
        .Formulas(69) = ""
        .Formulas(70) = ""
        .Formulas(71) = ""
        .Formulas(72) = ""
        .Formulas(73) = ""
        .Formulas(74) = ""
        .Formulas(75) = ""
        .Formulas(76) = ""
        .Formulas(77) = ""
        .Formulas(78) = ""
        .Formulas(79) = ""
        .Formulas(80) = ""
        .Formulas(81) = ""
        .Formulas(82) = ""
        .Formulas(83) = ""
        .Formulas(84) = ""
        .Formulas(85) = ""
        .Formulas(86) = ""
        .Formulas(87) = ""
        .Formulas(88) = ""
        .Formulas(89) = ""
        .Formulas(90) = ""
        .Formulas(91) = ""
        .Formulas(92) = ""
        .Formulas(93) = ""
        .Formulas(94) = ""
        .Formulas(95) = ""
        .Formulas(96) = ""
        .Formulas(97) = ""
        .Formulas(98) = ""
        .Formulas(99) = ""
        .Formulas(100) = ""
        .Formulas(101) = ""
        .Formulas(102) = ""
        .Formulas(103) = ""
        .Formulas(104) = ""
        .Formulas(105) = ""
        .Formulas(106) = ""
        .Formulas(107) = ""
        .Formulas(108) = ""
        .Formulas(109) = ""
        .Formulas(110) = ""
        .Formulas(111) = ""
        .Formulas(112) = ""
        .Formulas(113) = ""
        .Formulas(114) = ""
        .Formulas(115) = ""
        .Formulas(116) = ""
        .Formulas(117) = ""
        .Formulas(118) = ""
        .Formulas(119) = ""
        .Formulas(120) = ""
        .Formulas(121) = ""
        .Formulas(122) = ""
        .Formulas(123) = ""
        .Formulas(124) = ""
        .Formulas(125) = ""
        .Formulas(126) = ""
        .Formulas(127) = ""
        .Formulas(128) = ""
        .Formulas(129) = ""
        .Formulas(130) = ""
        .Formulas(131) = ""
        .Formulas(132) = ""
        .Formulas(133) = ""
        .Formulas(134) = ""
        .Formulas(135) = ""
        .Formulas(136) = ""
        .Formulas(137) = ""
        .Formulas(138) = ""
        .Formulas(139) = ""
        .Formulas(140) = ""
        .Formulas(141) = ""
        .Formulas(142) = ""
        .Formulas(143) = ""
        .Formulas(144) = ""
        .Formulas(145) = ""
        .Formulas(146) = ""
        .Formulas(147) = ""
        .Formulas(148) = ""
        .ReportFileName = Empty
    End With
    
End Sub
Sub Calcula_Porcentagem_Rec()
' Calcula a Porcentagem sobre o total de capas (Env/Mal) X qtde de capas por meia hora
    
Dim SomaQtd             As Long      '   Soma Quantidades
Dim SomaCapa            As Long      '   Soma Totais
Dim Porcentagem         As Integer   '   Contador de Loop
Dim SomaPorcentagem     As String    '   Divide a Soma das qtdes pela soma dos totais

    For Porcentagem = 1 To 48
        If Not FormataRelhoras.QtdEnv(Porcentagem) = "0" Or Not FormataRelhoras.QtdMal(Porcentagem) = "0" Then
            If Not Trim(FormataRelhoras.QtdEnv(Porcentagem)) = "" Or Not Trim(FormataRelhoras.QtdMal(Porcentagem)) = "" Then
                If Trim(FormataRelhoras.QtdEnv(Porcentagem)) = "" And Trim(FormataRelhoras.QtdMal(Porcentagem)) <> "" Then
                   FormataRelhoras.QtdEnv(Porcentagem) = "0"
                End If
                If Trim(FormataRelhoras.QtdEnv(Porcentagem)) <> "" And Trim(FormataRelhoras.QtdMal(Porcentagem)) = "" Then
                   FormataRelhoras.QtdMal(Porcentagem) = "0"
                End If
                SomaQtd = CLng(FormataRelhoras.QtdEnv(Porcentagem)) + CLng(FormataRelhoras.QtdMal(Porcentagem))
                SomaQtd = SomaQtd * 100
                SomaCapa = CLng(FormataRelhoras.SomaEnv) + CLng(FormataRelhoras.SomaMal)
                SomaPorcentagem = SomaQtd / SomaCapa
                FormataRelhoras.Porcentagem(Porcentagem) = StrAposVirgula(SomaPorcentagem)
            End If
        End If
    Next

End Sub
Function StrAposVirgula(StrPorcentagem As String)
' Esta função tem a finalidade de formatar e retornar a porcentagem com duas casas decimais
' sem arredondamento..

Dim ContaString    As Integer   'Conta o número de Caracteres da Porcentagem
Dim ContaQtdString As Integer   'Contador de Loop - para descobrir em que posição se encontra a vírgula
Dim StrVirgula     As String    'Variavel axiliar que guarda parte do texto
Dim StrNumVirgula  As String    'Numero da String dentro do texto

    ContaString = Len(StrPorcentagem)

    For ContaQtdString = 1 To ContaString
        
        StrVirgula = Mid$(StrPorcentagem, ContaQtdString, 1)
        
        If StrVirgula = "," Then
            StrNumVirgula = ContaQtdString
            ContaQtdString = ContaQtdString + 2
            StrAposVirgula = Mid$(StrPorcentagem, 1, ContaQtdString)
            Exit For
        Else
            StrAposVirgula = (StrPorcentagem)
        End If
        
    Next

End Function
Function TrataHoraInicial(Hora_Inicial As String) As String
    
    If Mid$(Hora_Inicial, 4, 2) = "" Then Exit Function
    
    If Mid$(Hora_Inicial, 4, 2) <= 30 Then
       TrataHoraInicial = Mid$(Hora_Inicial, 1, 2) & ":" & "00"
    Else
       TrataHoraInicial = Mid$(Hora_Inicial, 1, 2) & ":" & "30"
    End If

End Function
Function TrataHoraPeriodo(Hora_Periodo As String) As String

    If Mid$(Hora_Periodo, 4, 2) = "" Then Exit Function
    
    If Mid$(Hora_Periodo, 4, 2) <= 1 Then
       TrataHoraPeriodo = Mid$(Hora_Periodo, 1, 2) & ":" & "00"
    Else
       TrataHoraPeriodo = Mid$(Hora_Periodo, 1, 2) & ":" & "30"
    End If

End Function
Function Valores_Grade(HrInic, HrFinal As String, Qtde, TipoCapa As Integer) As String
'* Esta função retorna a quantidade de capas de Envelope e Malote para uma determina hora *'
    
Dim Periodo     As String
Dim Perinicial  As String
Dim PerFinal    As String
Dim QtdeCapa    As Integer
Dim CtaArrayFom As Integer
Dim Contador    As Integer
       
    Perinicial = HrInic  'Hora Inicial do Período
    PerFinal = HrFinal   'Hora Final   do Período
    QtdeCapa = Qtde      'Quantidade   de Capa - Malote / Envelope
     
    Perinicial = TrataHoraInicial(Perinicial)
    PerFinal = TrataHoraInicial(PerFinal)
    
    Select Case TipoCapa
    
        Case 1:
            'Tratamento de Envelope
            For CtaArrayFom = 1 To 48
                Periodo = Mid$(FormataRelhoras.Periodo(CtaArrayFom), 1, 5)
                
                If CtaArrayFom > 1 Then
                    Periodo = TrataHoraInicial(Periodo)
                End If
    
                If Periodo = Perinicial Or Periodo = PerFinal Then
                   FormataRelhoras.QtdEnv(CtaArrayFom) = Qtde
                   Exit Function
                End If
            Next
        
        Case 2:
            'Tratamento de Malote
            For CtaArrayFom = 1 To 48
                Periodo = Mid$(FormataRelhoras.Periodo(CtaArrayFom), 1, 5)
                If CtaArrayFom > 1 Then
                    Periodo = TrataHoraInicial(Periodo)
                End If
                If Periodo = Perinicial Or Periodo = PerFinal Then
                   FormataRelhoras.QtdMal(CtaArrayFom) = Qtde
                   Exit Function
                End If
            Next
    
    End Select
      
End Function
Sub OpcaoEnvelope()
'* - Traz a Linha de Envelope Linha 1 - *'
    
    GrdAcompRec.Column = 1
    FormataRelhoras.SomaEnv = 0
        
    For CountEnv = 1 To GrdAcompRec.RowCount
        GrdAcompRec.Row = CountEnv
        For Contador = 1 To 48
            horaInicial = Mid$(Horas.Periodo(Contador), 1, 5)
            horaFinal = Mid$(Horas.Periodo(Contador), 10, 5)
            
            If horaInicial = "" Then
               GrdAcompRec.Data = 0
               Exit For
            End If

            If GrdAcompRec.RowLabel = horaInicial Or GrdAcompRec.RowLabel = horaFinal And Horas.TipoEnv(Contador) = "E" Then
                GrdAcompRec.Data = Horas.Quantidade(Contador)
                Valores_Grade horaInicial, horaFinal, Horas.Quantidade(Contador), 1
                FormataRelhoras.SomaEnv = FormataRelhoras.SomaEnv + Horas.Quantidade(Contador)
                Exit For
            Else
                GrdAcompRec.Data = 0
            End If
        Next
    Next
    
End Sub
Sub OpcaoMalote()
'* - Traz a Linha de Malote   Linha 2 - *'
    
    GrdAcompRec.Column = 2
    FormataRelhoras.SomaMal = 0
    
    For CountMal = 1 To GrdAcompRec.RowCount
        GrdAcompRec.Row = CountMal
        For Contador = 1 To 48
            horaInicial = Mid$(Horas.Periodo(Contador), 1, 5)
            horaFinal = Mid$(Horas.Periodo(Contador), 10, 5)
    
            If horaInicial = "" Then
               GrdAcompRec.Data = 0
               Exit For
            End If
    
            If GrdAcompRec.RowLabel = horaInicial Or GrdAcompRec.RowLabel = horaFinal And Horas.TipoEnv(Contador) = "M" Then
                GrdAcompRec.Data = Horas.Quantidade(Contador)
                Valores_Grade horaInicial, horaFinal, Horas.Quantidade(Contador), 2
                FormataRelhoras.SomaMal = FormataRelhoras.SomaMal + Horas.Quantidade(Contador)
                Exit For
            Else
                GrdAcompRec.Data = 0
            End If
        Next
    Next
    
End Sub
Sub OpcaoTodos()
'* - Traz a linha de Todos    Linha 3 -*'
'* - A Linha 3 é o Total de Envelope + Malote

Dim SomaQtdes  As Integer

    GrdAcompRec.Column = 3

    For CountEnvMal = 1 To GrdAcompRec.RowCount
        GrdAcompRec.Row = CountEnvMal
        For Contador = 1 To 48
            horaInicial = Mid$(Horas.Periodo(Contador), 1, 5)
            horaFinal = Mid$(Horas.Periodo(Contador), 10, 5)
    
            If horaInicial = "" Then Exit For
    
            If GrdAcompRec.RowLabel = horaInicial Or GrdAcompRec.RowLabel = horaFinal Then
                SomaQtdes = SomaQtdes + Horas.Quantidade(Contador)
            End If
        Next
        GrdAcompRec.Data = SomaQtdes
        SomaQtdes = 0
    Next
        
End Sub
Function TrataRecepcao() As Boolean

    'Query que traz todas as horas e quantidade de Capas para data atual - Recepcionada
    Set qryAcompRec = Geral.Banco.CreateQuery("", "{Call RecHoraChegadaEnvMal(?)}")
    
    With qryAcompRec
        .rdoParameters(0) = Geral.DataProcessamento
        Set RsAcompRec = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsAcompRec.EOF Then
        GrdAcompRec.Visible = True
        cmdImprimir.Enabled = True
        horaInicial = TrataHoraInicial(RsAcompRec!MinHoraChegada)
    
        If Not IsNull(RsAcompRec!maxhoraCHegada) Then
            horaFinal = TrataHoraInicial(RsAcompRec!maxhoraCHegada)
            horaFinal = CalcSomaTime(horaFinal, "00:30")
        Else
            horaFinal = CalcSomaTime(horaInicial, "00:30")
        End If
    
        For CountCapas = 1 To RsAcompRec.RowCount
            '* Preenche type Horas com valores da Base de Dados *'
            Horas.Periodo(CountCapas) = RsAcompRec!Periodo
            Horas.PerForm1(CountCapas) = Mid$(RsAcompRec!Periodo, 1, 5)
            Horas.PerForm2(CountCapas) = Mid$(RsAcompRec!Periodo, 10, 5)
            Horas.Quantidade(CountCapas) = RsAcompRec!Envelope
            Horas.TipoEnv(CountCapas) = IIf(RsAcompRec!Tipo = "M", "M", "E") 'Trata (E)nvelope e (F)ininvest como Envelope
'            Horas.TipoEnv(CountCapas) = RsAcompRec!Tipo
            RsAcompRec.MoveNext
        Next
        
        TrataRecepcao = True
    Else

        TrataRecepcao = False
        GrdAcompRec.Visible = False
        cmdImprimir.Enabled = False

    End If

End Function
Sub TrataCaptura()

    'Query que traz todas as horas e quantidade de Capas para data atual - Capturada
    Set qryAcompCap = Geral.Banco.CreateQuery("", "{call Reccapturaenvmal(?)}")
    
    With qryAcompCap
        .rdoParameters(0) = Geral.DataProcessamento
        Set RsAcompCap = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsAcompCap.EOF Then
       
       hrFinalCap = RsAcompCap!MaxHoraCaptura
       hrFinalCap = CalcSomaTime(hrFinalCap, "01:00")
     
        For CtCapasCap = 1 To RsAcompCap.RowCount
        '* Preenche type Horas com valores da Base de Dados - Captura *'
            HorasCap.Periodo(CtCapasCap) = Mid$(RsAcompCap!Periodo, 10, 20)
            HorasCap.PerForm1(CtCapasCap) = Mid$(RsAcompCap!Periodo, 10, 5)
            HorasCap.PerForm2(CtCapasCap) = Mid$(RsAcompCap!Periodo, 19, 5)
            HorasCap.Quantidade(CtCapasCap) = RsAcompCap!Envelope
            HorasCap.TipoEnv(CtCapasCap) = IIf(RsAcompCap!Tipo = "M", "M", "E") 'Trata (E)nvelope e (F)ininvest como Envelope
            RsAcompCap.MoveNext
        Next
        
    Else
    
        hrFinalCap = horaFinal
    
    End If
        
End Sub
Sub OpcaoEnvelopeCaptura()
'* - Traz a Linha de Envelope Linha 4 - *'
    
    GrdAcompRec.Column = 4
    FormataRelhoras.SomaEnvCap = 0
        
    For CountEnvCap = 1 To GrdAcompRec.RowCount
        GrdAcompRec.Row = CountEnvCap
        For Contador = 1 To 48
            horaInicial = HorasCap.PerForm1(Contador)
            horaFinal = HorasCap.PerForm2(Contador)
            
            If horaInicial = "" Then
                GrdAcompRec.Data = 0
                Exit For
            End If
                
            If GrdAcompRec.RowLabel = horaInicial Or GrdAcompRec.RowLabel = horaFinal And HorasCap.TipoEnv(Contador) = "E" Then
                GrdAcompRec.Data = HorasCap.Quantidade(Contador)
                Valores_Grade_Cap horaInicial, horaFinal, HorasCap.Quantidade(Contador), 1
                FormataRelhoras.SomaEnvCap = FormataRelhoras.SomaEnvCap + HorasCap.Quantidade(Contador)
                Exit For
            Else
                GrdAcompRec.Data = 0
            End If
        Next
    Next
    
End Sub
Sub OpcaoMaloteCaptura()
'* - Traz a Linha de Malote  Linha 5 - *'
    
    GrdAcompRec.Column = 5
    FormataRelhoras.SomaMalCap = 0
    
    For CountMalCap = 1 To GrdAcompRec.RowCount
        GrdAcompRec.Row = CountMalCap
        For Contador = 1 To 48
            horaInicial = HorasCap.PerForm1(Contador)
            horaFinal = HorasCap.PerForm2(Contador)
    
            If horaInicial = "" Then
                GrdAcompRec.Data = 0
                Exit For
            End If
    
            If GrdAcompRec.RowLabel = horaInicial Or GrdAcompRec.RowLabel = horaFinal And HorasCap.TipoEnv(Contador) = "M" Then
                GrdAcompRec.Data = HorasCap.Quantidade(Contador)
                Valores_Grade_Cap horaInicial, horaFinal, HorasCap.Quantidade(Contador), 2
                FormataRelhoras.SomaMalCap = FormataRelhoras.SomaMalCap + HorasCap.Quantidade(Contador)
                Exit For
            Else
                GrdAcompRec.Data = 0
            End If
        Next
    Next
    
End Sub
Sub OpcaoTodosCaptura()
'* - Traz a linha de Todos    Linha 6 -*'
'* - A Linha 6 é o Total de Envelope Capturado + Malote Capturado

Dim SomaQtdes  As Integer

    GrdAcompRec.Column = 6

    For CtEnvMalCap = 1 To GrdAcompRec.RowCount
        GrdAcompRec.Row = CtEnvMalCap
        For Contador = 1 To 48
            horaInicial = HorasCap.PerForm1(Contador)
            horaFinal = HorasCap.PerForm2(Contador)
    
            If horaInicial = "" Then Exit For
    
            If GrdAcompRec.RowLabel = horaInicial Or GrdAcompRec.RowLabel = horaFinal Then
                SomaQtdes = SomaQtdes + HorasCap.Quantidade(Contador)
            End If
        Next
        GrdAcompRec.Data = SomaQtdes
        SomaQtdes = 0
    Next
        
End Sub
Sub Calcula_Porcentagem_Cap()
' Calcula a Porcentagem sobre o total de capas (Env/Mal) X qtde de capas por meia hora (Captura)
    
Dim SomaQtdCap             As Long      '   Soma Quantidades - Captura
Dim SomaCapaCap            As Long      '   Soma Totais      - Captura
Dim PorcentagemCap         As Integer   '   Contador de Loop - Captura
Dim SomaPorcentagemCap     As String    '   Divide a Soma das qtdes pela soma dos totais

    For PorcentagemCap = 1 To 48
        If Not FormataRelhoras.QtdEnvCap(PorcentagemCap) = "0" Or Not FormataRelhoras.QtdMalCap(PorcentagemCap) = "0" Then
            If Not Trim(FormataRelhoras.QtdEnvCap(PorcentagemCap)) = "" Or Not Trim(FormataRelhoras.QtdMalCap(PorcentagemCap)) = "" Then
                If Trim(FormataRelhoras.QtdEnvCap(PorcentagemCap)) = "" And Trim(FormataRelhoras.QtdMalCap(PorcentagemCap)) <> "" Then
                   FormataRelhoras.QtdEnvCap(PorcentagemCap) = "0"
                End If
                If Trim(FormataRelhoras.QtdEnvCap(PorcentagemCap)) <> "" And Trim(FormataRelhoras.QtdMalCap(PorcentagemCap)) = "" Then
                   FormataRelhoras.QtdMalCap(PorcentagemCap) = "0"
                End If
                SomaQtdCap = CLng(FormataRelhoras.QtdEnvCap(PorcentagemCap)) + CLng(FormataRelhoras.QtdMalCap(PorcentagemCap))
                SomaQtdCap = SomaQtdCap * 100
                SomaCapaCap = CLng(FormataRelhoras.SomaEnvCap) + CLng(FormataRelhoras.SomaMalCap)
                SomaPorcentagemCap = SomaQtdCap / SomaCapaCap
                FormataRelhoras.PorcentCap(PorcentagemCap) = StrAposVirgula(SomaPorcentagemCap)
            End If
        End If
    Next
End Sub
Sub ImprimeRecepcao()
'* Este botão aciona a impressão do Relatório de Recepção sendo
'  o parâmetro a data de processamento e detalhe de meia em meia hora *'

Dim CountLinhas As Integer

CountLinhas = 1
GrdAcompRec.Row = CountLinhas

    With RptAcompRec
    
        Screen.MousePointer = 1
        
            .WindowTop = 1
            .WindowLeft = 1
            .Connect = Geral.StringConexao
            .ReportFileName = App.path & "\AcompRecep.rpt"
            .WindowState = crptMaximized
            .WindowTitle = "Acompanhamento de Recepção de Capas - Envelope / Malote"
            .Formulas(0) = "HoraInicial = '" & GrdAcompRec.RowLabel & "'"
            .Formulas(1) = "dataprocessamento = '" & lblDataProc & "'"
            CountLinhas = GrdAcompRec.RowCount
            GrdAcompRec.Row = CountLinhas
            .Formulas(2) = "HoraFinal = '" & GrdAcompRec.RowLabel & "'"
            Call Prepara_Relatorio
            .Action = 0
            Call Limpa_Formula
        
        Screen.MousePointer = 0

    
    End With

End Sub
Sub ImprimeCaptura()
'* Este botão aciona a impressão do Relatório de Captura sendo
'  o parâmetro a data de processamento e detalhe de meia em meia hora *'

Dim CountLinhascap As Integer

CountLinhascap = 1
GrdAcompRec.Row = CountLinhascap

    With RptAcompRec
    
        Screen.MousePointer = 1
        
            .WindowTop = 1
            .WindowLeft = 1
            .Connect = Geral.StringConexao
            .ReportFileName = App.path & "\AcompCaptura.rpt"
            .WindowState = crptMaximized
            .WindowTitle = "Acompanhamento de Capturas de Capas - Envelope / Malote"
            .Formulas(0) = "HoraInicial = '" & GrdAcompRec.RowLabel & "'"
            .Formulas(1) = "dataprocessamento = '" & lblDataProc & "'"
            CountLinhascap = GrdAcompRec.RowCount
            GrdAcompRec.Row = CountLinhascap
            .Formulas(2) = "HoraFinal = '" & GrdAcompRec.RowLabel & "'"
            Call Prepara_Relatorio_cap
            .Action = 0
            Call Limpa_Formula
        
        Screen.MousePointer = 0

    
    End With

End Sub
Sub Prepara_Relatorio_cap()
'* Prepara o relátorio de recepção - formulas *'

    With RptAcompRec
    
    '---Formatação de Período
        .Formulas(3) = "Periodo1    = '" & FormataRelhoras.Periodo(1) & FormataRelhoras.Formata(1) & "'"
        .Formulas(4) = "Periodo2    = '" & FormataRelhoras.Periodo(2) & FormataRelhoras.Formata(2) & "'"
        .Formulas(5) = "Periodo3    = '" & FormataRelhoras.Periodo(3) & FormataRelhoras.Formata(3) & "'"
        .Formulas(6) = "Periodo4    = '" & FormataRelhoras.Periodo(4) & FormataRelhoras.Formata(4) & "'"
        .Formulas(7) = "Periodo5    = '" & FormataRelhoras.Periodo(5) & FormataRelhoras.Formata(5) & "'"
        .Formulas(8) = "Periodo6    = '" & FormataRelhoras.Periodo(6) & FormataRelhoras.Formata(6) & "'"
        .Formulas(9) = "Periodo7    = '" & FormataRelhoras.Periodo(7) & FormataRelhoras.Formata(7) & "'"
        .Formulas(10) = "Periodo8   = '" & FormataRelhoras.Periodo(8) & FormataRelhoras.Formata(8) & "'"
        .Formulas(11) = "Periodo9   = '" & FormataRelhoras.Periodo(9) & FormataRelhoras.Formata(9) & "'"
        .Formulas(12) = "Periodo10  = '" & FormataRelhoras.Periodo(10) & FormataRelhoras.Formata(10) & "'"
        .Formulas(13) = "Periodo11  = '" & FormataRelhoras.Periodo(11) & FormataRelhoras.Formata(11) & "'"
        .Formulas(14) = "Periodo12  = '" & FormataRelhoras.Periodo(12) & FormataRelhoras.Formata(12) & "'"
        .Formulas(15) = "Periodo13  = '" & FormataRelhoras.Periodo(13) & FormataRelhoras.Formata(13) & "'"
        .Formulas(16) = "Periodo14  = '" & FormataRelhoras.Periodo(14) & FormataRelhoras.Formata(14) & "'"
        .Formulas(17) = "Periodo15  = '" & FormataRelhoras.Periodo(15) & FormataRelhoras.Formata(15) & "'"
        .Formulas(18) = "Periodo16  = '" & FormataRelhoras.Periodo(16) & FormataRelhoras.Formata(16) & "'"
        .Formulas(19) = "Periodo17  = '" & FormataRelhoras.Periodo(17) & FormataRelhoras.Formata(17) & "'"
        .Formulas(20) = "Periodo18  = '" & FormataRelhoras.Periodo(18) & FormataRelhoras.Formata(18) & "'"
        .Formulas(21) = "Periodo19  = '" & FormataRelhoras.Periodo(19) & FormataRelhoras.Formata(19) & "'"
        .Formulas(22) = "Periodo20  = '" & FormataRelhoras.Periodo(20) & FormataRelhoras.Formata(20) & "'"
        .Formulas(23) = "Periodo21  = '" & FormataRelhoras.Periodo(21) & FormataRelhoras.Formata(21) & "'"
        .Formulas(24) = "Periodo22  = '" & FormataRelhoras.Periodo(22) & FormataRelhoras.Formata(22) & "'"
        .Formulas(25) = "Periodo23  = '" & FormataRelhoras.Periodo(23) & FormataRelhoras.Formata(23) & "'"
        .Formulas(26) = "Periodo24  = '" & FormataRelhoras.Periodo(24) & FormataRelhoras.Formata(24) & "'"
        .Formulas(27) = "Periodo25  = '" & FormataRelhoras.Periodo(25) & FormataRelhoras.Formata(25) & "'"
        .Formulas(28) = "Periodo26  = '" & FormataRelhoras.Periodo(26) & FormataRelhoras.Formata(26) & "'"
        .Formulas(29) = "Periodo27  = '" & FormataRelhoras.Periodo(27) & FormataRelhoras.Formata(27) & "'"
        .Formulas(30) = "Periodo28  = '" & FormataRelhoras.Periodo(28) & FormataRelhoras.Formata(28) & "'"
        .Formulas(31) = "Periodo29  = '" & FormataRelhoras.Periodo(29) & FormataRelhoras.Formata(29) & "'"
        .Formulas(32) = "Periodo30  = '" & FormataRelhoras.Periodo(30) & FormataRelhoras.Formata(30) & "'"
        .Formulas(33) = "Periodo31  = '" & FormataRelhoras.Periodo(31) & FormataRelhoras.Formata(31) & "'"
        .Formulas(96) = "Periodo32  = '" & FormataRelhoras.Periodo(32) & FormataRelhoras.Formata(32) & "'"
        .Formulas(97) = "Periodo33  = '" & FormataRelhoras.Periodo(33) & FormataRelhoras.Formata(33) & "'"
        .Formulas(98) = "Periodo34  = '" & FormataRelhoras.Periodo(34) & FormataRelhoras.Formata(34) & "'"
        .Formulas(99) = "Periodo35  = '" & FormataRelhoras.Periodo(35) & FormataRelhoras.Formata(35) & "'"
        .Formulas(100) = "Periodo36 = '" & FormataRelhoras.Periodo(36) & FormataRelhoras.Formata(36) & "'"
        .Formulas(101) = "Periodo37 = '" & FormataRelhoras.Periodo(37) & FormataRelhoras.Formata(37) & "'"
        .Formulas(102) = "Periodo38 = '" & FormataRelhoras.Periodo(38) & FormataRelhoras.Formata(38) & "'"
        .Formulas(103) = "Periodo39 = '" & FormataRelhoras.Periodo(39) & FormataRelhoras.Formata(39) & "'"
        .Formulas(104) = "Periodo40 = '" & FormataRelhoras.Periodo(40) & FormataRelhoras.Formata(40) & "'"
        .Formulas(105) = "Periodo41 = '" & FormataRelhoras.Periodo(41) & FormataRelhoras.Formata(41) & "'"
        .Formulas(106) = "Periodo42 = '" & FormataRelhoras.Periodo(42) & FormataRelhoras.Formata(42) & "'"
        .Formulas(107) = "Periodo43 = '" & FormataRelhoras.Periodo(43) & FormataRelhoras.Formata(43) & "'"
        .Formulas(108) = "Periodo44 = '" & FormataRelhoras.Periodo(44) & FormataRelhoras.Formata(44) & "'"
        .Formulas(109) = "Periodo45 = '" & FormataRelhoras.Periodo(45) & FormataRelhoras.Formata(45) & "'"
        .Formulas(110) = "Periodo46 = '" & FormataRelhoras.Periodo(46) & FormataRelhoras.Formata(46) & "'"
        .Formulas(111) = "Periodo47 = '" & FormataRelhoras.Periodo(47) & FormataRelhoras.Formata(47) & "'"
        .Formulas(112) = "Periodo48 = '" & FormataRelhoras.Periodo(48) & FormataRelhoras.Formata(48) & "'"
     
     '---Formatação de Quantide de Envelopes
        .Formulas(34) = "Envelope1  = '" & FormataRelhoras.QtdEnvCap(1) & "'"
        .Formulas(35) = "Envelope2  = '" & FormataRelhoras.QtdEnvCap(2) & "'"
        .Formulas(36) = "Envelope3  = '" & FormataRelhoras.QtdEnvCap(3) & "'"
        .Formulas(37) = "Envelope4  = '" & FormataRelhoras.QtdEnvCap(4) & "'"
        .Formulas(38) = "Envelope5  = '" & FormataRelhoras.QtdEnvCap(5) & "'"
        .Formulas(39) = "Envelope6  = '" & FormataRelhoras.QtdEnvCap(6) & "'"
        .Formulas(40) = "Envelope7  = '" & FormataRelhoras.QtdEnvCap(7) & "'"
        .Formulas(41) = "Envelope8  = '" & FormataRelhoras.QtdEnvCap(8) & "'"
        .Formulas(42) = "Envelope9  = '" & FormataRelhoras.QtdEnvCap(9) & "'"
        .Formulas(43) = "Envelope10 = '" & FormataRelhoras.QtdEnvCap(10) & "'"
        .Formulas(44) = "Envelope11 = '" & FormataRelhoras.QtdEnvCap(11) & "'"
        .Formulas(45) = "Envelope12 = '" & FormataRelhoras.QtdEnvCap(12) & "'"
        .Formulas(46) = "Envelope13 = '" & FormataRelhoras.QtdEnvCap(13) & "'"
        .Formulas(47) = "Envelope14 = '" & FormataRelhoras.QtdEnvCap(14) & "'"
        .Formulas(48) = "Envelope15 = '" & FormataRelhoras.QtdEnvCap(15) & "'"
        .Formulas(49) = "Envelope16 = '" & FormataRelhoras.QtdEnvCap(16) & "'"
        .Formulas(50) = "Envelope17 = '" & FormataRelhoras.QtdEnvCap(17) & "'"
        .Formulas(51) = "Envelope18 = '" & FormataRelhoras.QtdEnvCap(18) & "'"
        .Formulas(52) = "Envelope19 = '" & FormataRelhoras.QtdEnvCap(19) & "'"
        .Formulas(53) = "Envelope20 = '" & FormataRelhoras.QtdEnvCap(20) & "'"
        .Formulas(54) = "Envelope21 = '" & FormataRelhoras.QtdEnvCap(21) & "'"
        .Formulas(55) = "Envelope22 = '" & FormataRelhoras.QtdEnvCap(22) & "'"
        .Formulas(56) = "Envelope23 = '" & FormataRelhoras.QtdEnvCap(23) & "'"
        .Formulas(57) = "Envelope24 = '" & FormataRelhoras.QtdEnvCap(24) & "'"
        .Formulas(58) = "Envelope25 = '" & FormataRelhoras.QtdEnvCap(25) & "'"
        .Formulas(59) = "Envelope26 = '" & FormataRelhoras.QtdEnvCap(26) & "'"
        .Formulas(60) = "Envelope27 = '" & FormataRelhoras.QtdEnvCap(27) & "'"
        .Formulas(61) = "Envelope28 = '" & FormataRelhoras.QtdEnvCap(28) & "'"
        .Formulas(62) = "Envelope29 = '" & FormataRelhoras.QtdEnvCap(29) & "'"
        .Formulas(63) = "Envelope30 = '" & FormataRelhoras.QtdEnvCap(30) & "'"
        .Formulas(113) = "Envelope31 = '" & FormataRelhoras.QtdEnvCap(31) & "'"
        .Formulas(114) = "Envelope32 = '" & FormataRelhoras.QtdEnvCap(32) & "'"
        .Formulas(115) = "Envelope33 = '" & FormataRelhoras.QtdEnvCap(33) & "'"
        .Formulas(116) = "Envelope34 = '" & FormataRelhoras.QtdEnvCap(34) & "'"
        .Formulas(117) = "Envelope35 = '" & FormataRelhoras.QtdEnvCap(35) & "'"
        .Formulas(118) = "Envelope36 = '" & FormataRelhoras.QtdEnvCap(36) & "'"
        .Formulas(119) = "Envelope37 = '" & FormataRelhoras.QtdEnvCap(37) & "'"
        .Formulas(120) = "Envelope38 = '" & FormataRelhoras.QtdEnvCap(38) & "'"
        .Formulas(121) = "Envelope39 = '" & FormataRelhoras.QtdEnvCap(39) & "'"
        .Formulas(122) = "Envelope40 = '" & FormataRelhoras.QtdEnvCap(40) & "'"
        .Formulas(123) = "Envelope41 = '" & FormataRelhoras.QtdEnvCap(41) & "'"
        .Formulas(124) = "Envelope42 = '" & FormataRelhoras.QtdEnvCap(42) & "'"
        .Formulas(125) = "Envelope43 = '" & FormataRelhoras.QtdEnvCap(43) & "'"
        .Formulas(126) = "Envelope44 = '" & FormataRelhoras.QtdEnvCap(44) & "'"
        .Formulas(127) = "Envelope45 = '" & FormataRelhoras.QtdEnvCap(45) & "'"
        .Formulas(128) = "Envelope46 = '" & FormataRelhoras.QtdEnvCap(46) & "'"
        .Formulas(129) = "Envelope47 = '" & FormataRelhoras.QtdEnvCap(47) & "'"
        .Formulas(130) = "Envelope48 = '" & FormataRelhoras.QtdEnvCap(48) & "'"
    
    '---Formatação de Quantidades de Malotes
        .Formulas(65) = "Malote1   = '" & FormataRelhoras.QtdMalCap(1) & "'"
        .Formulas(66) = "Malote2   = '" & FormataRelhoras.QtdMalCap(2) & "'"
        .Formulas(67) = "Malote3   = '" & FormataRelhoras.QtdMalCap(3) & "'"
        .Formulas(68) = "Malote4   = '" & FormataRelhoras.QtdMalCap(4) & "'"
        .Formulas(69) = "Malote5   = '" & FormataRelhoras.QtdMalCap(5) & "'"
        .Formulas(70) = "Malote6   = '" & FormataRelhoras.QtdMalCap(6) & "'"
        .Formulas(71) = "Malote7   = '" & FormataRelhoras.QtdMalCap(7) & "'"
        .Formulas(72) = "Malote8   = '" & FormataRelhoras.QtdMalCap(8) & "'"
        .Formulas(73) = "Malote9   = '" & FormataRelhoras.QtdMalCap(9) & "'"
        .Formulas(74) = "Malote10  = '" & FormataRelhoras.QtdMalCap(10) & "'"
        .Formulas(75) = "Malote11  = '" & FormataRelhoras.QtdMalCap(11) & "'"
        .Formulas(76) = "Malote12  = '" & FormataRelhoras.QtdMalCap(12) & "'"
        .Formulas(77) = "Malote13  = '" & FormataRelhoras.QtdMalCap(13) & "'"
        .Formulas(78) = "Malote14  = '" & FormataRelhoras.QtdMalCap(14) & "'"
        .Formulas(79) = "Malote15  = '" & FormataRelhoras.QtdMalCap(15) & "'"
        .Formulas(80) = "Malote16  = '" & FormataRelhoras.QtdMalCap(16) & "'"
        .Formulas(81) = "Malote17  = '" & FormataRelhoras.QtdMalCap(17) & "'"
        .Formulas(82) = "Malote18  = '" & FormataRelhoras.QtdMalCap(18) & "'"
        .Formulas(83) = "Malote19  = '" & FormataRelhoras.QtdMalCap(19) & "'"
        .Formulas(84) = "Malote20  = '" & FormataRelhoras.QtdMalCap(20) & "'"
        .Formulas(85) = "Malote21  = '" & FormataRelhoras.QtdMalCap(21) & "'"
        .Formulas(86) = "Malote22  = '" & FormataRelhoras.QtdMalCap(22) & "'"
        .Formulas(87) = "Malote23  = '" & FormataRelhoras.QtdMalCap(23) & "'"
        .Formulas(88) = "Malote24  = '" & FormataRelhoras.QtdMalCap(24) & "'"
        .Formulas(89) = "Malote25  = '" & FormataRelhoras.QtdMalCap(25) & "'"
        .Formulas(90) = "Malote26  = '" & FormataRelhoras.QtdMalCap(26) & "'"
        .Formulas(91) = "Malote27  = '" & FormataRelhoras.QtdMalCap(27) & "'"
        .Formulas(92) = "Malote28  = '" & FormataRelhoras.QtdMalCap(28) & "'"
        .Formulas(93) = "Malote29  = '" & FormataRelhoras.QtdMalCap(29) & "'"
        .Formulas(94) = "Malote30  = '" & FormataRelhoras.QtdMalCap(30) & "'"
        .Formulas(131) = "Malote31 = '" & FormataRelhoras.QtdMalCap(31) & "'"
        .Formulas(132) = "Malote32 = '" & FormataRelhoras.QtdMalCap(32) & "'"
        .Formulas(133) = "Malote33 = '" & FormataRelhoras.QtdMalCap(33) & "'"
        .Formulas(134) = "Malote34 = '" & FormataRelhoras.QtdMalCap(34) & "'"
        .Formulas(135) = "Malote35 = '" & FormataRelhoras.QtdMalCap(35) & "'"
        .Formulas(136) = "Malote36 = '" & FormataRelhoras.QtdMalCap(36) & "'"
        .Formulas(137) = "Malote37 = '" & FormataRelhoras.QtdMalCap(37) & "'"
        .Formulas(138) = "Malote38 = '" & FormataRelhoras.QtdMalCap(38) & "'"
        .Formulas(139) = "Malote39 = '" & FormataRelhoras.QtdMalCap(39) & "'"
        .Formulas(140) = "Malote40 = '" & FormataRelhoras.QtdMalCap(40) & "'"
        .Formulas(141) = "Malote41 = '" & FormataRelhoras.QtdMalCap(41) & "'"
        .Formulas(142) = "Malote42 = '" & FormataRelhoras.QtdMalCap(42) & "'"
        .Formulas(143) = "Malote43 = '" & FormataRelhoras.QtdMalCap(43) & "'"
        .Formulas(144) = "Malote44 = '" & FormataRelhoras.QtdMalCap(44) & "'"
        .Formulas(145) = "Malote45 = '" & FormataRelhoras.QtdMalCap(45) & "'"
        .Formulas(146) = "Malote46 = '" & FormataRelhoras.QtdMalCap(46) & "'"
        .Formulas(147) = "Malote47 = '" & FormataRelhoras.QtdMalCap(47) & "'"
        .Formulas(148) = "Malote48 = '" & FormataRelhoras.QtdMalCap(48) & "'"
      
      '--Total de Capas
        .Formulas(149) = "TotalEnvelope = '" & FormataRelhoras.SomaEnvCap & "'"
        .Formulas(150) = "TotalMalote= '" & FormataRelhoras.SomaMalCap & "'"
      
      '--Porcentagem
        .Formulas(151) = "Porcentagem1 = '" & FormataRelhoras.PorcentCap(1) & "'"
        .Formulas(152) = "Porcentagem2 = '" & FormataRelhoras.PorcentCap(2) & "'"
        .Formulas(153) = "Porcentagem3 = '" & FormataRelhoras.PorcentCap(3) & "'"
        .Formulas(154) = "Porcentagem4 = '" & FormataRelhoras.PorcentCap(4) & "'"
        .Formulas(155) = "Porcentagem5 = '" & FormataRelhoras.PorcentCap(5) & "'"
        .Formulas(156) = "Porcentagem6 = '" & FormataRelhoras.PorcentCap(6) & "'"
        .Formulas(157) = "Porcentagem7 = '" & FormataRelhoras.PorcentCap(7) & "'"
        .Formulas(158) = "Porcentagem8 = '" & FormataRelhoras.PorcentCap(8) & "'"
        .Formulas(159) = "Porcentagem9 = '" & FormataRelhoras.PorcentCap(9) & "'"
        .Formulas(160) = "Porcentagem10= '" & FormataRelhoras.PorcentCap(10) & "'"
        .Formulas(161) = "Porcentagem11= '" & FormataRelhoras.PorcentCap(11) & "'"
        .Formulas(162) = "Porcentagem12= '" & FormataRelhoras.PorcentCap(12) & "'"
        .Formulas(163) = "Porcentagem13= '" & FormataRelhoras.PorcentCap(13) & "'"
        .Formulas(164) = "Porcentagem14= '" & FormataRelhoras.PorcentCap(14) & "'"
        .Formulas(165) = "Porcentagem15= '" & FormataRelhoras.PorcentCap(15) & "'"
        .Formulas(166) = "Porcentagem16= '" & FormataRelhoras.PorcentCap(16) & "'"
        .Formulas(167) = "Porcentagem17= '" & FormataRelhoras.PorcentCap(17) & "'"
        .Formulas(168) = "Porcentagem18= '" & FormataRelhoras.PorcentCap(18) & "'"
        .Formulas(169) = "Porcentagem19= '" & FormataRelhoras.PorcentCap(19) & "'"
        .Formulas(171) = "Porcentagem20= '" & FormataRelhoras.PorcentCap(20) & "'"
        .Formulas(172) = "Porcentagem21= '" & FormataRelhoras.PorcentCap(21) & "'"
        .Formulas(173) = "Porcentagem22= '" & FormataRelhoras.PorcentCap(22) & "'"
        .Formulas(174) = "Porcentagem23= '" & FormataRelhoras.PorcentCap(23) & "'"
        .Formulas(175) = "Porcentagem24= '" & FormataRelhoras.PorcentCap(24) & "'"
        .Formulas(176) = "Porcentagem25= '" & FormataRelhoras.PorcentCap(25) & "'"
        .Formulas(177) = "Porcentagem26= '" & FormataRelhoras.PorcentCap(26) & "'"
        .Formulas(178) = "Porcentagem27= '" & FormataRelhoras.PorcentCap(27) & "'"
        .Formulas(179) = "Porcentagem28= '" & FormataRelhoras.PorcentCap(28) & "'"
        .Formulas(181) = "Porcentagem29= '" & FormataRelhoras.PorcentCap(29) & "'"
        .Formulas(182) = "Porcentagem30= '" & FormataRelhoras.PorcentCap(30) & "'"
        .Formulas(183) = "Porcentagem31= '" & FormataRelhoras.PorcentCap(31) & "'"
        .Formulas(184) = "Porcentagem32= '" & FormataRelhoras.PorcentCap(32) & "'"
        .Formulas(185) = "Porcentagem33= '" & FormataRelhoras.PorcentCap(33) & "'"
        .Formulas(186) = "Porcentagem34= '" & FormataRelhoras.PorcentCap(34) & "'"
        .Formulas(187) = "Porcentagem35= '" & FormataRelhoras.PorcentCap(35) & "'"
        .Formulas(188) = "Porcentagem36= '" & FormataRelhoras.PorcentCap(36) & "'"
        .Formulas(189) = "Porcentagem37= '" & FormataRelhoras.PorcentCap(37) & "'"
        .Formulas(191) = "Porcentagem38= '" & FormataRelhoras.PorcentCap(38) & "'"
        .Formulas(192) = "Porcentagem39= '" & FormataRelhoras.PorcentCap(39) & "'"
        .Formulas(193) = "Porcentagem40= '" & FormataRelhoras.PorcentCap(40) & "'"
        .Formulas(194) = "Porcentagem41= '" & FormataRelhoras.PorcentCap(41) & "'"
        .Formulas(195) = "Porcentagem42= '" & FormataRelhoras.PorcentCap(42) & "'"
        .Formulas(196) = "Porcentagem43= '" & FormataRelhoras.PorcentCap(43) & "'"
        .Formulas(197) = "Porcentagem44= '" & FormataRelhoras.PorcentCap(44) & "'"
        .Formulas(198) = "Porcentagem45= '" & FormataRelhoras.PorcentCap(45) & "'"
        .Formulas(199) = "Porcentagem46= '" & FormataRelhoras.PorcentCap(46) & "'"
        .Formulas(200) = "Porcentagem47= '" & FormataRelhoras.PorcentCap(47) & "'"
        .Formulas(201) = "Porcentagem48= '" & FormataRelhoras.PorcentCap(48) & "'"
        .Formulas(202) = "AgProcessadora= '" & Geral.AgenciaCentral & "'"
    End With

End Sub
Function Valores_Grade_Cap(HrInic, HrFinal As String, Qtde, TipoCapa As Integer) As String
'* Esta função retorna a quantidade de capas de Envelope e Malote para uma determina hora *'
    
Dim PeriodoCap     As String
Dim PerinicialCap  As String
Dim PerFinalCap    As String
Dim QtdeCapaCap    As Integer
Dim CtaArrayFomCap As Integer
Dim ContadorCap    As Integer
       
    PerinicialCap = HrInic  'Hora Inicial do Período
    PerFinalCap = HrFinal   'Hora Final   do Período
    QtdeCapaCap = Qtde      'Quantidade   de Capa - Malote / Envelope
     
    PerinicialCap = TrataHoraInicial(PerinicialCap)
    PerFinalCap = TrataHoraInicial(PerFinalCap)
    
    Select Case TipoCapa
    
        Case 1:
            'Tratamento de Envelope
            For CtaArrayFomCap = 1 To 48
                PeriodoCap = Mid$(FormataRelhoras.Periodo(CtaArrayFomCap), 1, 5)
                
                If CtaArrayFomCap > 1 Then
                    PeriodoCap = TrataHoraInicial(PeriodoCap)
                End If
    
                If PeriodoCap = PerinicialCap Or PeriodoCap = PerFinalCap Then
                   FormataRelhoras.QtdEnvCap(CtaArrayFomCap) = Qtde
                   Exit Function
                End If
            Next
        
        Case 2:
            'Tratamento de Malote
            For CtaArrayFomCap = 1 To 48
                PeriodoCap = Mid$(FormataRelhoras.Periodo(CtaArrayFomCap), 1, 5)
                If CtaArrayFomCap > 1 Then
                    PeriodoCap = TrataHoraInicial(PeriodoCap)
                End If
                If PeriodoCap = PerinicialCap Or PeriodoCap = PerFinalCap Then
                   FormataRelhoras.QtdMalCap(CtaArrayFomCap) = Qtde
                   Exit Function
                End If
            Next
    
    End Select
      
End Function

