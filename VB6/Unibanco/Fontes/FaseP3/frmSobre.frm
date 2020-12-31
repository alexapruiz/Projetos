VERSION 5.00
Begin VB.Form frmSobre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre o Sistema"
   ClientHeight    =   2985
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   732
      Left            =   120
      Picture         =   "frmSobre.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   372
      Left            =   4320
      TabIndex        =   0
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   3492
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Multi-Documentos por Imagem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   3492
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   3492
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   3492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   1440
      X2              =   4920
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

'Picture1.Picture = LoadResPicture(4, vbResBitmap)
Label1.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
Label2.Caption = "MDI - Unibanco"
Label4.Caption = "Copyright(C) Centro de Competência Backoffice"

End Sub

