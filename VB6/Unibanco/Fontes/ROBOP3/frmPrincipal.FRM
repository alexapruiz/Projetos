VERSION 5.00
Begin VB.Form frmPrincipal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   Caption         =   "MDI  Unibanco  -  MultiAgência  - Caixa Robô"
   ClientHeight    =   8448
   ClientLeft      =   1368
   ClientTop       =   1680
   ClientWidth     =   11652
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8448
   ScaleWidth      =   11652
   WindowState     =   2  'Maximized
   Begin VB.Label LblMSG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde.. Copiando Log's para o Servidor"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   20.4
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   2490
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   6555
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Caption = "MDI Unibanco - Caixa Robô - " & App.Major & App.Minor & App.Revision
End Sub
