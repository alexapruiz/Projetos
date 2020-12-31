VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form AlteraDataMovimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alterar data de movimento"
   ClientHeight    =   1248
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   4092
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1248
   ScaleWidth      =   4092
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   888
      Left            =   192
      TabIndex        =   3
      Top             =   144
      Width           =   2544
      Begin DATEEDITLib.DateEdit txtDataProcessamento 
         Height          =   384
         Left            =   384
         TabIndex        =   0
         Top             =   336
         Width           =   1776
         _Version        =   65537
         _ExtentX        =   3133
         _ExtentY        =   677
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2952
      TabIndex        =   1
      Top             =   240
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2952
      TabIndex        =   2
      Top             =   672
      Width           =   972
   End
End
Attribute VB_Name = "AlteraDataMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bOk       As Boolean
Private m_sData     As String
Public Function ShowModal(ByRef pDataMovimento As String) As Boolean


    Me.txtDataProcessamento.Text = Format(Format(Geral.DataProcessamento, "0000/00/00"), "dd/mm/yyyy")

    Me.Show vbModal
    
    pDataMovimento = m_sData
    
    ShowModal = m_bOk
    
    

End Function

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
End Sub

Private Sub CmdOK_Click()

    m_bOk = False
    
    If (Trim(txtDataProcessamento.Text) = "") Then
'        If Not IsDate(Format(txtDataProcessamento.Text, "00/00/0000")) Then
        MsgBox "Data inválida.", vbExclamation
        SelecionarTexto txtDataProcessamento
        txtDataProcessamento.SetFocus
        Exit Sub
'        End If
    End If
    
    m_bOk = True
    
    m_sData = Format(Format(txtDataProcessamento.Text, "00/00/0000"), "yyyymmdd")
    
    Unload Me
End Sub


Private Sub txtDataProcessamento_GotFocus()
    SelecionarTexto txtDataProcessamento
End Sub


Private Sub txtDataProcessamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(Format(txtDataProcessamento.Text, "00/00/0000")) Then
            CmdOK_Click
        End If
    End If
End Sub

