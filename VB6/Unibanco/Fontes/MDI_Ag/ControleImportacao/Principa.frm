VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Principal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8544
   ClientLeft      =   48
   ClientTop       =   636
   ClientWidth     =   12216
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8544
   ScaleWidth      =   12216
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar BarMain 
      Align           =   2  'Align Bottom
      Height          =   324
      Left            =   0
      TabIndex        =   0
      Top             =   8220
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   572
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17622
            MinWidth        =   17622
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "05/12/00"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "23:18"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuImportar 
      Caption         =   "&Importar"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Dim i As Long
    Dim iTam As Long
    
    Dim y As Single
    
    iTam = ScaleHeight / 256
    
    y = 0
    
    For i = 0 To 255
        Line (0, y)-(ScaleWidth, y + iTam), RGB(i, i, i), BF
        y = y + iTam
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Geral.Banco.Close
End Sub

Private Sub mnuImportar_Click()
    Importacao.Show vbModal
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub
