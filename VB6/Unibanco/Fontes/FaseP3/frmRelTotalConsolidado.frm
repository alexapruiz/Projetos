VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRelTotalConsolidado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Total Consolidado"
   ClientHeight    =   3924
   ClientLeft      =   3012
   ClientTop       =   2208
   ClientWidth     =   4812
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3924
   ScaleWidth      =   4812
   Begin VB.Frame fraGravar 
      Caption         =   "Gravar Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3012
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   4092
      Begin VB.TextBox txtArquivo 
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   14
         Top             =   720
         Width           =   2820
      End
      Begin VB.CommandButton cmdDiretorio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3564
         Picture         =   "frmRelTotalConsolidado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Width           =   348
      End
      Begin VB.TextBox txtDiretorio 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   288
         Left            =   120
         TabIndex        =   12
         Top             =   1536
         Width           =   3828
      End
      Begin VB.CommandButton cmdArqGravar 
         Caption         =   "&Gravar Arquivo"
         Height          =   312
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   1572
      End
      Begin VB.CommandButton cmdArqCancelar 
         Caption         =   "&Sair"
         Height          =   312
         Left            =   2160
         TabIndex        =   10
         Top             =   2280
         Width           =   1572
      End
      Begin VB.Label lblArquivo 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   504
      End
      Begin VB.Label lblDiretorio 
         AutoSize        =   -1  'True
         Caption         =   "Diretório"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   732
      End
   End
   Begin ComctlLib.ProgressBar pgbProcesso 
      Height          =   132
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   4332
      _ExtentX        =   7641
      _ExtentY        =   233
      _Version        =   327682
      Appearance      =   0
   End
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
      Begin VB.CommandButton cmdGravarArquivo 
         Cancel          =   -1  'True
         Caption         =   "&Gerar Arquivo"
         Height          =   312
         Left            =   1680
         TabIndex        =   8
         Top             =   3000
         Width           =   1212
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   312
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancel 
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
Attribute VB_Name = "frmRelTotalConsolidado"
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
        cmdImp.Enabled = False
        cmdGravarArquivo.Enabled = False
    End If
    
End Sub

Private Sub cmdArqCancelar_Click()

    fraGravar.Visible = False
    fraPrincipal.Enabled = True
    txtArquivo.Text = ""
    txtDiretorio.Text = ""
    
End Sub

Private Sub cmdArqGravar_Click()

Dim bExisteArquivo As Boolean, sArquivo As String

On Error GoTo Err_cmdArqGravar

    If Trim(txtArquivo.Text) = "" Then
          Beep
          MsgBox "Favor informar o nome do arquivo para gravar os dados do relatório", vbInformation, Me.Caption
          txtArquivo.SetFocus
          Exit Sub
     End If
    
     If Trim(txtDiretorio) = "" Then
          Beep
          MsgBox "Favor informar o diretório para gravar os dados do relatório", vbInformation, Me.Caption
          cmdDiretorio.SetFocus
          Exit Sub
     End If
    
     sArquivo = frmRelTotalConsolidado.txtDiretorio & "\" & Trim(frmRelTotalConsolidado.txtArquivo.Text) & ".csv"
     
    bExisteArquivo = True
    Open sArquivo For Input As #1
    Close #1
    If bExisteArquivo Then
        If MsgBox("Arquivo já existe, deseja apagá-lo", vbQuestion + vbYesNo, App.Title) = vbNo Then
            txtArquivo.SetFocus
            Exit Sub
        End If
    End If
     
     Kill (sArquivo)

    If RelTotalConsolidado.RelEstatisticaConsolidado Then
        MsgBox "Finalizada a gravação das informações do relatório em arquivo", vbInformation + vbOKOnly, App.Title
        cmdArqCancelar_Click
    End If
    Exit Sub
    
Err_cmdArqGravar:

    If Err.Number = 53 Then
        bExisteArquivo = False
        Resume Next
    End If
    
    Beep
    Close #1
    MsgBox "Não foi possível gravar o arquivo." & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDiretorio_Click()

     txtDiretorio.Text = BrowseForFolder(Me, 17)

End Sub

Private Sub cmdGravarArquivo_Click()

    fraPrincipal.Enabled = False
    fraGravar.Visible = True
    txtArquivo.SetFocus
    
End Sub

Private Sub cmdImp_Click()
    
    Call RelTotalConsolidado.RelEstatisticaConsolidado
    
End Sub

Private Sub Form_Activate()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    fraGravar.Visible = False
    lstDias.Enabled = False
    cmdImp.Enabled = False
    cmdGravarArquivo.Enabled = False
    pgbProcesso.Visible = False
    
    Call CarregaCombo

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

Private Sub Form_Unload(Cancel As Integer)

    Set qryDiasMovto = Nothing
    Set rsDiasMovto = Nothing
    Set qryMeses = Nothing
    Set rsMeses = Nothing

End Sub

Private Sub lstDias_Click()

    If lstDias.SelCount > 0 Then
        cmdImp.Enabled = True
        cmdGravarArquivo.Enabled = True
    Else
        cmdImp.Enabled = False
        cmdGravarArquivo.Enabled = False
    End If
    
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
