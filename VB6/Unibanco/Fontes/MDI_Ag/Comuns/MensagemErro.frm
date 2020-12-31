VERSION 5.00
Begin VB.Form MensagemErro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   108
   ClientTop       =   540
   ClientWidth     =   7824
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7824
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRetorno 
      Caption         =   "Cancelar"
      Height          =   372
      Index           =   1
      Left            =   3906
      TabIndex        =   3
      Top             =   1800
      Width           =   1032
   End
   Begin VB.CommandButton cmdDetalhe 
      Caption         =   "Detalhes"
      Height          =   372
      Left            =   6780
      TabIndex        =   2
      Top             =   1800
      Width           =   972
   End
   Begin VB.CommandButton cmdRetorno 
      Caption         =   "Repetir"
      Height          =   372
      Index           =   0
      Left            =   2886
      TabIndex        =   1
      Top             =   1800
      Width           =   1032
   End
   Begin VB.Label lblErro 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   7700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aviso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7820
   End
End
Attribute VB_Name = "MensagemErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Texto As String
Public Erro As String
Public ErroBanco As String
Public Retorno As Byte
Public Sub Mostrar()
    lblErro = vbCrLf & Texto
End Sub

Private Sub cmdDetalhe_Click()
    Dim sMens As String
    
    sMens = ""
    
    If Len(Trim(Erro)) > 0 Then
        sMens = sMens & "Mensagem do VB" & vbCrLf & Erro & vbCrLf & vbCrLf
    End If
    
    If Len(Trim(ErroBanco)) > 0 Then
        sMens = sMens & "Mensagem do Banco de Dados" & vbCrLf & ErroBanco
    End If
    
    MsgBox sMens, vbInformation + vbOKOnly, "Detalhes do Aviso"
End Sub

Private Sub cmdRetorno_Click(Index As Integer)
    Retorno = Index
    Me.Hide
End Sub


