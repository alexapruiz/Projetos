VERSION 5.00
Begin VB.Form FrmCalcDig 
   Caption         =   "Calcula CMC7"
   ClientHeight    =   2700
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check_Top 
      Caption         =   "&Always on Top"
      Height          =   192
      Left            =   2448
      TabIndex        =   21
      Top             =   2412
      Width           =   1395
   End
   Begin VB.CommandButton CmdClipBoard 
      Caption         =   "Send to ClipBoard"
      Height          =   300
      Left            =   4824
      TabIndex        =   20
      Top             =   2376
      Width           =   1560
   End
   Begin VB.TextBox Text_Campo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4650
      TabIndex        =   11
      Top             =   1440
      Width           =   1380
   End
   Begin VB.TextBox Text_Campo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   3420
      TabIndex        =   10
      Top             =   1440
      Width           =   1155
   End
   Begin VB.TextBox Text_Campo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2340
      TabIndex        =   9
      Top             =   1440
      Width           =   1020
   End
   Begin VB.TextBox Text_Digito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1815
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1404
      Width           =   228
   End
   Begin VB.TextBox Text_Banco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1008
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "409"
      Top             =   576
      Width           =   435
   End
   Begin VB.TextBox Text_Agencia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1008
      MaxLength       =   4
      TabIndex        =   2
      Top             =   972
      Width           =   570
   End
   Begin VB.TextBox Text_Conta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1008
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1404
      Width           =   750
   End
   Begin VB.TextBox Text_Cheque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1008
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1836
      Width           =   750
   End
   Begin VB.TextBox Text_tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1008
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "5"
      Top             =   2268
      Width           =   228
   End
   Begin VB.TextBox Text_Comp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1008
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "018"
      Top             =   180
      Width           =   408
   End
   Begin VB.CommandButton Cmd_Calcula 
      Caption         =   "Calcula"
      Height          =   552
      Left            =   3564
      TabIndex        =   8
      Top             =   360
      Width           =   1164
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Campo3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4965
      TabIndex        =   19
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Campo2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   18
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2535
      TabIndex        =   17
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   252
      TabIndex        =   16
      Top             =   612
      Width           =   588
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   252
      TabIndex        =   15
      Top             =   1044
      Width           =   744
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   252
      TabIndex        =   14
      Top             =   1476
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cheque:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   252
      TabIndex        =   13
      Top             =   1908
      Width           =   696
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   288
      TabIndex        =   12
      Top             =   2340
      Width           =   444
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   252
      TabIndex        =   7
      Top             =   252
      Width           =   492
   End
End
Attribute VB_Name = "FrmCalcDig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Top_Click()
    If Check_Top Then
        SetTopWindow Me.hWnd, True
       
    Else
        SetTopWindow Me.hWnd, False
       
    End If
End Sub

Private Sub Cmd_Calcula_Click()
    Dim Calc As New CalculoCheque
    
    Calc.Comp = Text_Comp
    Calc.Banco = Text_Banco
    Calc.Agencia = Text_Agencia
    Calc.Conta = Text_Conta & Text_Digito
    Calc.NumeroCheque = Text_Cheque
    Calc.Tipificacao = Text_tipo
    
    Call Calc.Calcula
    
    Text_Campo1 = Calc.Campo1
    Text_Campo2 = Calc.Campo2
    Text_Campo3 = Calc.Campo3
    
    Text_Cheque.Text = Format(Val(Text_Cheque.Text) + 1, "000000")

End Sub

Private Sub CmdClipBoard_Click()
    If (Text_Campo1 & Text_Campo2 & Text_Campo3) <> "" Then
        'Clipboard.SetText "'" & Text_Campo1 & Text_Campo2 & Text_Campo3 & "'"
        Clipboard.Clear
        Clipboard.SetText Text_Campo1 & Text_Campo2 & Text_Campo3
    End If
End Sub
Private Sub Text_Agencia_Change()
    If Len(Text_Agencia.Text) = Text_Agencia.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_Agencia_GotFocus()
    Text_Agencia.SelStart = 0
    Text_Agencia.SelLength = Len(Text_Agencia.Text)
End Sub
Private Sub Text_Agencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Text_Banco_Change()
    If Len(Text_Banco.Text) = Text_Banco.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_Banco_GotFocus()
    Text_Banco.SelStart = 0
    Text_Banco.SelLength = Len(Text_Banco.Text)
End Sub
Private Sub Text_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Text_Campo1_GotFocus()
    SelecionarTexto Text_Campo1
    
    Clipboard.Clear
    
    Clipboard.SetText Text_Campo1.Text
End Sub


Private Sub Text_Campo2_GotFocus()
    SelecionarTexto Text_Campo2
    
    Clipboard.Clear
    
    Clipboard.SetText Text_Campo2.Text

End Sub


Private Sub Text_Campo3_GotFocus()

    SelecionarTexto Text_Campo3
    
    Clipboard.Clear
    
    Clipboard.SetText Text_Campo3.Text

End Sub


Private Sub Text_Cheque_Change()
    If Len(Text_Cheque.Text) = Text_Cheque.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_Cheque_GotFocus()
    Text_Cheque.SelStart = 0
    Text_Cheque.SelLength = Len(Text_Cheque.Text)
End Sub
Private Sub Text_Comp_Change()
    If Len(Text_Comp.Text) = Text_Comp.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_Comp_GotFocus()
    Text_Comp.SelStart = 0
    Text_Comp.SelLength = Len(Text_Comp.Text)
End Sub
Private Sub Text_Comp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Text_Conta_Change()
    If Len(Text_Conta.Text) = Text_Conta.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_Conta_GotFocus()
    Text_Conta.SelStart = 0
    Text_Conta.SelLength = Len(Text_Conta.Text)
End Sub
Private Sub Text_Conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Text_Digito_Change()
    If Len(Text_Digito.Text) = Text_Digito.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_Digito_GotFocus()
    Text_Digito.SelStart = 0
    Text_Digito.SelLength = Len(Text_Digito.Text)
End Sub
Private Sub Text_Digito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Text_tipo_Change()
    If Len(Text_tipo.Text) = Text_tipo.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text_tipo_GotFocus()
    Text_tipo.SelStart = 0
    Text_tipo.SelLength = Len(Text_tipo.Text)
End Sub
Private Sub Text_tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
