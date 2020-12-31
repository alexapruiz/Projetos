VERSION 5.00
Begin VB.Form Central 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdChamar 
      Caption         =   "Chamada"
      Height          =   405
      Left            =   2430
      TabIndex        =   6
      Top             =   2130
      Width           =   945
   End
   Begin VB.CommandButton CmdDescer 
      Caption         =   "Descer"
      Height          =   405
      Left            =   2430
      TabIndex        =   5
      Top             =   930
      Width           =   915
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Subir"
      Height          =   375
      Left            =   2430
      TabIndex        =   4
      Top             =   450
      Width           =   915
   End
   Begin VB.PictureBox p2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   330
      ScaleHeight     =   405
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   870
      Width           =   495
   End
   Begin VB.PictureBox p1 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   330
      ScaleHeight     =   405
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1380
      Width           =   495
   End
   Begin VB.PictureBox p0 
      BackColor       =   &H0000FFFF&
      Height          =   465
      Left            =   330
      ScaleHeight     =   405
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   1890
      Width           =   495
   End
   Begin VB.PictureBox p3 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   330
      ScaleHeight     =   405
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Central"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Elev As New Elevador
Private Sub Cmd1_Click()

    Select Case Elev.Andar
    Case 0
        Call Elev.Subir(p0, p1)
    Case 1
        Call Elev.Subir(p1, p2)
    Case 2
        Call Elev.Subir(p2, p3)
    Case 3
        MsgBox "Não é possível subir"
    End Select
End Sub

Private Sub CmdChamar_Click()

    Dim x As Integer

    'Simular que o elevador foi chamado pelo terceiro andar estando no andar terreo
    Call Sleep(2000)
    Call Elev.Subir(p0, p1)
    Call Sleep(2000)
    Call Elev.Subir(p1, p2)
    Call Sleep(2000)
    Call Elev.Subir(p2, p3)
End Sub
Private Sub CmdDescer_Click()

    Select Case Elev.Andar
    Case 0
        MsgBox "Não é possível Descer"
    Case 1
        Call Elev.Descer(p1, p0)
    Case 2
        Call Elev.Descer(p2, p1)
    Case 3
        Call Elev.Descer(p3, p2)
    End Select
End Sub
