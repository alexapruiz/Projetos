VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4092
   ClientLeft      =   48
   ClientTop       =   300
   ClientWidth     =   5808
   LinkTopic       =   "Form1"
   ScaleHeight     =   4092
   ScaleWidth      =   5808
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4464
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abort Shutdown"
      Height          =   396
      Left            =   1968
      TabIndex        =   11
      Top             =   3024
      Width           =   1404
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1344
      TabIndex        =   10
      Text            =   "0"
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1344
      TabIndex        =   8
      Text            =   "0"
      Top             =   1344
      Width           =   972
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1344
      TabIndex        =   3
      Text            =   "15"
      Top             =   1008
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1344
      TabIndex        =   2
      Text            =   "Agora vamos desligar..."
      Top             =   672
      Width           =   2076
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1344
      TabIndex        =   1
      Text            =   "\\Estacao43"
      Top             =   336
      Width           =   2076
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shutdown"
      Height          =   396
      Left            =   1968
      TabIndex        =   0
      Top             =   2544
      Width           =   1404
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   528
      TabIndex        =   12
      Top             =   3696
      Width           =   4620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Reboot"
      Height          =   192
      Left            =   144
      TabIndex        =   9
      Top             =   1776
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Force"
      Height          =   192
      Left            =   144
      TabIndex        =   7
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Time Out"
      Height          =   192
      Left            =   96
      TabIndex        =   6
      Top             =   1104
      Width           =   648
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   192
      Left            =   96
      TabIndex        =   5
      Top             =   720
      Width           =   816
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estação"
      Height          =   192
      Left            =   96
      TabIndex        =   4
      Top             =   384
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long

Private Sub Command1_Click()
    Dim MName   As String
    Dim msg     As String
    Dim tOut    As Long
    Dim Force   As Long
    Dim Reboot  As Long
    Dim ret     As Long
    
'    MName = "\\Estacao43" & Chr(0)
'    msg = "Agora vou desligar..." & Chr(0)
'    tOut = 15
'    Force = 0
'    Reboot = 1
    
    MName = Text1.Text & Chr(0)
    msg = Text2.Text & Chr(0)
    tOut = Val(Text3.Text)
    Force = Val(Text4.Text)
    Reboot = Val(Text5.Text)

    ret = InitiateSystemShutdown(MName, msg, tOut, Force, Reboot)
    
    Timer1.Enabled = False
    If ret Then
        Timer1.Enabled = True
    End If
    
    Label6.Caption = "InitiateSystemShutdown retornou " & ret
    
    
    
End Sub

Private Sub Command2_Click()

    Dim MName       As String
    
    MName = Text1.Text & Chr(0)
    Timer1.Enabled = False
    
    Label6.Caption = "AbortSystemShutdown retornou " & AbortSystemShutdown(MName)
    
End Sub


Private Sub Timer1_Timer()
    Text3.Text = Val(Text3.Text) - 1
End Sub


