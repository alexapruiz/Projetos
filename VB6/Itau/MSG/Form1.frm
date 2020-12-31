VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mensagem"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdEnviar5 
      Caption         =   "Enviar 5"
      Height          =   345
      Left            =   1140
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox CboDestinatario 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   210
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   3165
   End
   Begin VB.CommandButton CmdEnviar 
      Caption         =   "Enviar"
      Height          =   345
      Left            =   3300
      TabIndex        =   3
      Top             =   2640
      Width           =   945
   End
   Begin VB.TextBox TxtMensagem 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1410
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Destinatário"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1170
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ENVIAR As Integer
Private Sub CmdEnviar_Click()

    Dim x As Double

    'Verificar se foi selecionado um destinatário
    If CboDestinatario.ListIndex = -1 Then
        MsgBox "Informe o Destinatário"
        CboDestinatario.SetFocus
        Exit Sub
    End If

    'Verificar se foi digitada uma mensagem
    If Len(Trim(TxtMensagem.Text)) = 0 Then
        MsgBox "Digite uma mensagem para '" & CboDestinatario.Text & "'"
    End If

    'Envia a mensagem
    x = Shell("SEND '" & TxtMensagem.Text & "' to " & CboDestinatario.Text, vbHide)
End Sub
Private Sub CmdEnviar5_Click()

    'Verificar se foi selecionado um destinatário
'    If CboDestinatario.ListIndex = -1 Then
'        MsgBox "Informe o Destinatário"
'        CboDestinatario.SetFocus
'        Exit Sub
'    End If

    'Verificar se foi digitada uma mensagem
'    If Len(Trim(TxtMensagem.Text)) = 0 Then
'        MsgBox "Digite uma mensagem para '" & CboDestinatario.Text & "'"
'    End If

'    ENVIAR = 5
'    Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()

'    Dim x As Double
'
'    If ENVIAR > 0 Then
'        x = Shell("SEND '" & TxtMensagem.Text & "' to " & CboDestinatario.Text, vbHide)
'        ENVIAR = ENVIAR - 1
'    End If
'
'    If ENVIAR = 0 Then
'        Timer1.Enabled = False
'    End If
End Sub

