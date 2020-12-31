VERSION 5.00
Begin VB.Form MotivoExclusao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motivo da Exclusão"
   ClientHeight    =   3324
   ClientLeft      =   1464
   ClientTop       =   1368
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   324
      Left            =   3300
      TabIndex        =   2
      Top             =   2904
      Width           =   1092
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   324
      Left            =   2100
      TabIndex        =   1
      Top             =   2904
      Width           =   1092
   End
   Begin VB.TextBox TxtMotivo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2028
      Left            =   108
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   768
      Width           =   6384
   End
   Begin VB.PictureBox Picture2 
      Height          =   264
      Left            =   2040
      ScaleHeight     =   216
      ScaleWidth      =   1716
      TabIndex        =   7
      Top             =   108
      Width           =   1764
      Begin VB.Label LblValorEnv_Mal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   228
         Left            =   96
         TabIndex        =   8
         Top             =   0
         Width           =   1560
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   120
      ScaleHeight     =   216
      ScaleWidth      =   1824
      TabIndex        =   5
      Top             =   108
      Width           =   1872
      Begin VB.Label LblNroEnv_Mal 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro. "
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
         Height          =   228
         Left            =   48
         TabIndex        =   6
         Top             =   0
         Width           =   1764
      End
   End
   Begin VB.PictureBox PctMalote 
      Height          =   264
      Left            =   108
      ScaleHeight     =   216
      ScaleWidth      =   1836
      TabIndex        =   3
      Top             =   468
      Width           =   1884
      Begin VB.Label Label11 
         Caption         =   "Motivo da Exclusão"
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
         Height          =   228
         Left            =   36
         TabIndex        =   4
         Top             =   0
         Width           =   1752
      End
   End
   Begin VB.Label lblEnvFininvest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   3936
      TabIndex        =   9
      Top             =   96
      Width           =   2424
   End
End
Attribute VB_Name = "MotivoExclusao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Result As Boolean
Private qryInsereMotivoExclusao As rdoQuery
Private qryAtualizaStatusCapa As rdoQuery
Private qryGetMotivoExclusao As rdoQuery
Private qryRemoveMotivoExclusao As rdoQuery
Private qryAtualizaStatusDocumentosCapa As rdoQuery

Private Sub CmdCancelar_Click()
  Result = False
  Me.Hide

End Sub
Private Sub CmdOK_Click()

  On Error GoTo ErroMotivo

  Dim RsMotExc As rdoResultset
  Dim sSql As String

  Result = False

  'Verificar se o IdCapa foi passado corretamente
  If Val(LblValorEnv_Mal.Tag) = 0 Then
    MsgBox "Não foi possível ler a identificação da capa.", vbInformation + vbOKOnly, App.Title '
    Exit Sub
  End If

  'Verificar se foi informado um Motivo de Exclusão
  If Len(Trim(TxtMotivo.Text)) = 0 Then
    MsgBox "Nenhum Motivo foi Informado.", vbInformation, App.Title
    TxtMotivo.SetFocus
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass

  'Verificar se esta Capa já foi excluída
  sSql = Geral.DataProcessamento & " , " & Val(LblValorEnv_Mal.Tag)

  Set qryGetMotivoExclusao = Geral.Banco.CreateQuery("", "{call GetMotivoExclusao (" & sSql & ")}")

  Set RsMotExc = qryGetMotivoExclusao.OpenResultset(rdOpenStatic, rdConcurReadOnly)

  If Not RsMotExc.EOF Then
    'Já Existe - Excluir Motivo Antigo
    Set qryRemoveMotivoExclusao = Geral.Banco.CreateQuery("", "{? = call RemoveMotivoExclusao (?,?)}")
    With qryRemoveMotivoExclusao
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento  'Data Proc.
      .rdoParameters(2) = Val(LblValorEnv_Mal.Tag) 'IdCapa
      .Execute
    End With

    If qryRemoveMotivoExclusao(0).Value = 1 Then
      MsgBox "Ocorreu um erro ao excluir motivo de exclusão antigo.", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If
  End If

  'Gravar Motivo de Exclusão da Capa
  Set qryInsereMotivoExclusao = Geral.Banco.CreateQuery("", "{? = call InsereMotivoExclusao (?,?,?)}")
  With qryInsereMotivoExclusao
    .rdoParameters(0).Direction = rdParamReturnValue
    .rdoParameters(1) = Geral.DataProcessamento  'Data Proc.
    .rdoParameters(2) = Val(LblValorEnv_Mal.Tag) 'IdCapa
    .rdoParameters(3) = UCase(TxtMotivo.Text)    'MotivoExclusao
    .Execute
  End With

  If qryInsereMotivoExclusao(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao inserir motivo de exclusão.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  'Atualizar Status da Capa para 'D'
  Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusCapa (?,?,?)}")
  With qryAtualizaStatusCapa
    .rdoParameters(1) = Geral.DataProcessamento    'Data Proc.
    .rdoParameters(2) = Val(LblValorEnv_Mal.Tag)   'IdCapa
    .rdoParameters(3) = "D"                        'Status
    .Execute
  End With

  If qryAtualizaStatusCapa(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao atualizar status da capa.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  'Atualizar status e a ocorrencia de todos os documentos da capa
  Set qryAtualizaStatusDocumentosCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusDocumentosCapa (?,?,?,?)}")
  With qryAtualizaStatusDocumentosCapa
      .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Output
      .rdoParameters(1) = Geral.DataProcessamento          'Data Proc.
      .rdoParameters(2) = Val(LblValorEnv_Mal.Tag)         'IdCapa
      .rdoParameters(3) = "D"                              'Status
      .rdoParameters(4) = 999                              'Ocorrencia
      .Execute
  End With

  If qryAtualizaStatusDocumentosCapa(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao atualizar status dos documentos.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  Screen.MousePointer = vbDefault

  Result = True
  On Error GoTo 0
  Me.Hide
  Exit Sub

ErroMotivo:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro na exclusão do Envelope/Malote.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
  End Select
End Sub

Private Sub Form_Activate()
    
    'Apresenta Label informando Envelope Fininvest
    If Len(MotivoExclusao.LblValorEnv_Mal) = 10 Or Len(MotivoExclusao.LblValorEnv_Mal) = 9 Then
        lblEnvFininvest = "Envelope Fininvest"
    Else
        lblEnvFininvest = ""
    End If

End Sub

