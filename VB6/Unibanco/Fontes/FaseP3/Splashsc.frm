VERSION 5.00
Begin VB.Form SplashScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   3564
   ClientLeft      =   1488
   ClientTop       =   2580
   ClientWidth     =   6072
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3564
   ScaleWidth      =   6072
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3432
      Left            =   60
      ScaleHeight     =   3384
      ScaleWidth      =   5904
      TabIndex        =   0
      Top             =   60
      Width           =   5952
      Begin VB.Label Mensagem 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   1560
         TabIndex        =   1
         Top             =   1500
         Width           =   3216
      End
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
'   Left = (Screen.Width - ScaleWidth) / 2
'   Top = (Screen.Height - ScaleHeight) / 2
End Sub


Private Sub Label1_Click()

End Sub

