VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   2655
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1620
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   540
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sortear"
      Height          =   510
      Left            =   3960
      TabIndex        =   0
      Top             =   1350
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    For x = 1 To 1000
        x = Right(Now, 1)

        If x > 4 Then x = x - 5
        If Len(Text1(x).Text) > 0 Then
            MsgBox Text1(x)
            Exit For
        End If
    Next x
End Sub
