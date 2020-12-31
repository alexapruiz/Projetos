VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form DocumentoNaoDefinido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documento não definido"
   ClientHeight    =   2268
   ClientLeft      =   3780
   ClientTop       =   1488
   ClientWidth     =   3516
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2268
   ScaleWidth      =   3516
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   312
      Left            =   444
      TabIndex        =   4
      Top             =   1824
      Width           =   1068
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   312
      Left            =   1824
      TabIndex        =   3
      Top             =   1824
      Width           =   1068
   End
   Begin VB.Frame Frame1 
      Height          =   1656
      Left            =   144
      TabIndex        =   0
      Top             =   72
      Width           =   3168
      Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
         Height          =   420
         Left            =   312
         TabIndex        =   1
         Top             =   696
         Width           =   2484
         _Version        =   65537
         _ExtentX        =   4382
         _ExtentY        =   741
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   11
         BackColor       =   -2147483643
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   300
         TabIndex        =   2
         Top             =   456
         Width           =   468
      End
   End
End
Attribute VB_Name = "DocumentoNaoDefinido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InformouValor As Boolean
Public Valor As Currency
Private Sub CmdCancelar_Click()

  InformouValor = False
  Me.Hide
End Sub

Private Sub CmdOK_Click()

  'Verificar se foi informado um valor
  If Val(TxtValor.Text) <> 0 Then
    Valor = Val(TxtValor.Text) / 100
    InformouValor = True
    Me.Hide
  Else
    MsgBox "Nenhum Valor Informado.", vbInformation, App.Title
    TxtValor.SetFocus
  End If
End Sub
Private Sub Label1_Click()

End Sub


Private Sub Form_Activate()

  Me.Left = (Screen.Width - Me.Width) / 2
End Sub

