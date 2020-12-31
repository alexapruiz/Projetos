VERSION 5.00
Begin VB.Form frmComplRegOcorr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complemento de Ocorrência"
   ClientHeight    =   1380
   ClientLeft      =   1164
   ClientTop       =   1608
   ClientWidth     =   10068
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   10068
   Begin VB.CommandButton cmdOkComplemento 
      Caption         =   "&Ok"
      Height          =   300
      Left            =   8160
      TabIndex        =   1
      Top             =   240
      Width           =   1692
   End
   Begin VB.TextBox txtComplOcorrencia 
      Height          =   348
      Left            =   192
      MaxLength       =   80
      TabIndex        =   0
      Top             =   720
      Width           =   9708
   End
   Begin VB.Label lblComplemento 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   192
      Left            =   204
      TabIndex        =   2
      Top             =   528
      Width           =   864
   End
End
Attribute VB_Name = "frmComplRegOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Descricao As String

Private Sub cmdOkComplemento_Click()
    
    m_Descricao = Trim(txtComplOcorrencia.Text)
    Unload Me

End Sub

Private Sub Form_Activate()

    'Centraliza o form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    txtComplOcorrencia.SelStart = 0
    txtComplOcorrencia.SelLength = txtComplOcorrencia.MaxLength
    txtComplOcorrencia.SetFocus
    
End Sub

Private Sub Form_Load()

    If Len(m_Descricao) > 0 Then m_Descricao = Trim(m_Descricao)
    
    txtComplOcorrencia = m_Descricao
    
End Sub

Private Sub txtComplOcorrencia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then cmdOkComplemento.SetFocus
    'Se cancelou então retorna descrição de entrada
    If KeyAscii = vbKeyEscape Then
        txtComplOcorrencia.Text = m_Descricao
        cmdOkComplemento_Click
    End If

End Sub
