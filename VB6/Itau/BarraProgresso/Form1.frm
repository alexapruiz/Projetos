VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   2145
   WindowState     =   1  'Minimized
   Begin VB.CommandButton CmdAtualizar 
      Caption         =   "Pause"
      Height          =   315
      Left            =   1290
      TabIndex        =   9
      Top             =   1260
      Width           =   735
   End
   Begin VB.TextBox TxtFim 
      Height          =   285
      Left            =   630
      TabIndex        =   8
      Text            =   "17:30"
      Top             =   1260
      Width           =   585
   End
   Begin VB.TextBox TxtInicio 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Text            =   "08:30"
      Top             =   1260
      Width           =   555
   End
   Begin VB.PictureBox pic4 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3090
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2100
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pic3 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2460
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pic2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1800
      Picture         =   "Form1.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   2070
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pic1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1200
      Picture         =   "Form1.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2070
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   270
      Top             =   2040
   End
   Begin Threed.SSPanel p 
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   1890
      _Version        =   65536
      _ExtentX        =   3334
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodType       =   1
      Font3D          =   1
   End
   Begin VB.Label Label2 
      Caption         =   "v."
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   960
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   1050
      TabIndex        =   6
      Top             =   900
      Width           =   975
   End
   Begin VB.Label LblPercent 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAtualizar_Click()

    If UCase(CmdAtualizar.Caption) = "PAUSE" Then
        Timer1.Enabled = False
        CmdAtualizar.Caption = "INICIAR"
    Else
        Timer1.Enabled = True
        CmdAtualizar.Caption = "Pause"
    End If
End Sub

Private Sub Form_Load()

    Label2.Caption = "v." & App.Major & "." & App.Minor & "." & App.Revision
    Call Atualiza
End Sub
Private Sub Timer1_Timer()

    Call Atualiza
End Sub

Private Sub Atualiza()
    Dim Ini         As String
    Dim Fim         As String
    Dim Diferenca   As Long
    Dim Total       As Long

    'Define Hora Inicial
    Ini = TxtInicio.Text

    'Define Hora Final
    Fim = TxtFim.Text

    'Calcula percentual total
    Total = DateDiff("s", Ini, Fim)

    Diferenca = DateDiff("s", Ini, Format(Now, "hh:mm:ss"))

    If (Diferenca / Total) * 100 >= 100 Then
        p.FloodPercent = 100
    Else
        p.FloodPercent = (Diferenca / Total) * 100
    End If

    LblPercent.Caption = Format((Diferenca / Total) * 100, "00.0000")
    
    Me.Caption = LblPercent.Caption

    'Label regressivo
    Label1.Caption = DateDiff("s", Format(Now, "hh:mm:ss"), TxtFim.Text)

    pic1.Visible = False
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False

    If Val(p.FloodPercent) < 30 Then
        pic1.Visible = True
    ElseIf Val(p.FloodPercent) > 30 And Val(p.FloodPercent) < 50 Then
        pic2.Visible = True
    ElseIf Val(p.FloodPercent) > 50 And Val(p.FloodPercent) < 80 Then
        pic3.Visible = True
    ElseIf Val(p.FloodPercent) > 80 Then
        pic4.Visible = True
    End If
End Sub
