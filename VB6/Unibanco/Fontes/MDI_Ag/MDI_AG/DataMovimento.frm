VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form DataMovimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data do Movimento"
   ClientHeight    =   1500
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2153
      TabIndex        =   2
      Top             =   1008
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   833
      TabIndex        =   1
      Top             =   1008
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   120
      TabIndex        =   3
      Top             =   24
      Width           =   3972
      Begin DATEEDITLib.DateEdit dtDataMovimento 
         Height          =   372
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   1452
         _Version        =   65537
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   93
         Text            =   "09112000"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "09112000"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data do Movimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   276
         Width           =   2100
      End
   End
End
Attribute VB_Name = "DataMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_bOk               As Boolean
Dim m_lProcessamento    As Long
Public Function ShowModal(ByRef pDataProcessamento As Long) As Boolean

    Me.Show vbModal
    
    If m_bOk Then
        pDataProcessamento = m_lProcessamento
        ShowModal = True
    Else
        ShowModal = False
    End If

End Function


Private Sub btnCancel_Click()
    m_bOk = False
    
    Unload Me
End Sub

Private Sub btnOK_Click()
    m_bOk = True
    m_lProcessamento = CLng(dtDataMovimento.InverseText)
    
    Unload Me
End Sub


Private Sub dtDataMovimento_GotFocus()
    dtDataMovimento.SelStart = 0
    dtDataMovimento.SelLength = Len(dtDataMovimento.Text)
    dtDataMovimento.SetFocus
End Sub


Private Sub Form_Activate()

    dtDataMovimento.Text = Right(Geral.DataProcessamento, 2) & _
        Mid(Geral.DataProcessamento, 5, 2) & _
        Left(Geral.DataProcessamento, 4)

End Sub

