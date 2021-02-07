VERSION 5.00
Begin VB.Form Calendario 
   Caption         =   "SGB - Calendário"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSelecionar 
      Caption         =   "&Selecionar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3315
      Width           =   1095
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sai&r"
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   3315
      Width           =   1095
   End
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DataSelecionada As String
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub CmdSelecionar_Click()

    DataSelecionada = c.Value
    
    Me.Hide
End Sub
Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
