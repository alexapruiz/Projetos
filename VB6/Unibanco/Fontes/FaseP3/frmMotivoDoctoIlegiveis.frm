VERSION 5.00
Begin VB.Form frmMotivoDoctoIlegiveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documento Ilegível"
   ClientHeight    =   2304
   ClientLeft      =   2028
   ClientTop       =   3324
   ClientWidth     =   6456
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2304
   ScaleWidth      =   6456
   Begin VB.Frame fraPrincipal 
      Caption         =   " Motivos "
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
      Width           =   5988
      Begin VB.ComboBox cmbMotivos 
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
         Left            =   384
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   576
         Width           =   5196
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   312
         Left            =   1368
         TabIndex        =   1
         Top             =   1296
         Width           =   1572
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   312
         Left            =   3144
         TabIndex        =   2
         Top             =   1296
         Width           =   1572
      End
   End
End
Attribute VB_Name = "frmMotivoDoctoIlegiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_CodigoMotivo As Long

Private Sub cmbMotivos_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cmdOk_Click
    If KeyAscii = vbKeyEscape Then cmdCancelar_Click
    
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    If cmbMotivos.ListIndex = -1 Then
        MsgBox "Favor selecionar um dos motivos!", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    'Carrega em variavel o código do motivo de docto ilegível
    m_CodigoMotivo = cmbMotivos.ItemData(cmbMotivos.ListIndex)
    
    Unload Me
    
End Sub

Private Sub CarregaCombo()
    
Dim qryMotivoIlegiveis  As rdoQuery
Dim rsMotivos           As rdoResultset
    
On Error GoTo Err_CarregaCombo
    
    Screen.MousePointer = vbHourglass
    
    Set qryMotivoIlegiveis = Geral.Banco.CreateQuery("", "{call  GetMotivoIlegiveis(" & Null & ")}")
    Set rsMotivos = qryMotivoIlegiveis.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    Do Until rsMotivos.EOF()
        cmbMotivos.AddItem rsMotivos!Descricao
        cmbMotivos.ItemData(cmbMotivos.NewIndex) = rsMotivos!CodigoMotivo
        rsMotivos.MoveNext
    Loop

Exit_CarregaCombo:
    
    Screen.MousePointer = vbDefault
    Set rsMotivos = Nothing
    If Not (qryMotivoIlegiveis Is Nothing) Then qryMotivoIlegiveis.Close
    
    Exit Sub

Err_CarregaCombo:
    
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na leitura dos código de motivos para documento ilegível.", Err, rdoErrors)
        Case vbCancel, vbRetry
    End Select
    GoTo Exit_CarregaCombo

End Sub

Private Sub Form_Load()


    Call CarregaCombo
    
    Me.Top = DocumentoDesconhecido.Top
    Me.Left = (Screen.Width - Me.Width) / 2
    
    m_CodigoMotivo = 0
    
    If cmbMotivos.ListCount > 0 Then SendKeys "{F4}"
    
End Sub


