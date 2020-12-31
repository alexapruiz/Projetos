VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmImpressao 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox MyImp 
      Height          =   612
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   1080
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmImpressao.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H80000009&
      Height          =   3972
      Left            =   120
      ScaleHeight     =   3924
      ScaleWidth      =   2604
      TabIndex        =   0
      Top             =   120
      Width           =   2652
   End
End
Attribute VB_Name = "FrmImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    MyImp.Print "Teste Impressão"
    
End Sub
