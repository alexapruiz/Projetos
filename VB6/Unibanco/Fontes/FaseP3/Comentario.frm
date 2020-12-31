VERSION 5.00
Begin VB.Form Comentario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comentários"
   ClientHeight    =   2040
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   4776
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4776
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   396
      Left            =   3684
      TabIndex        =   2
      Top             =   684
      Width           =   972
   End
   Begin VB.TextBox txtComentario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   132
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   168
      Width           =   3384
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   396
      Left            =   3684
      TabIndex        =   1
      Top             =   168
      Width           =   972
   End
End
Attribute VB_Name = "Comentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_sStr          As String
Dim m_Ok            As Boolean
Public Function ShowModal(ByRef pComentario As String) As Boolean

    If Trim(pComentario) <> "" Then
        txtComentario.Text = Trim(pComentario)
        SelecionarTexto txtComentario
    End If

    Me.Show vbModal
    
    ShowModal = m_Ok
    
    pComentario = m_sStr

End Function


Private Sub cmdCancelar_Click()

    m_sStr = ""
    m_Ok = False
    Unload Me
End Sub

Private Sub cmdOk_Click()

    m_sStr = txtComentario.Text
    m_Ok = True
    Unload Me

End Sub


