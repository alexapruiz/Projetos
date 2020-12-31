VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FusaoDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fusão  Automática"
   ClientHeight    =   1680
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   7536
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   7536
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFusao 
      Height          =   330
      Left            =   6960
      Picture         =   "FusaoDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   495
      Width           =   375
   End
   Begin MSComDlg.CommonDialog DlgFusao 
      Left            =   8064
      Top             =   672
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4001
      TabIndex        =   1
      Top             =   1128
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2321
      TabIndex        =   0
      Top             =   1128
      Width           =   1215
   End
   Begin VB.TextBox TxtFusao 
      Height          =   370
      Left            =   2400
      MaxLength       =   60
      TabIndex        =   3
      Top             =   470
      Width           =   4932
   End
   Begin VB.Label LblFusao 
      Caption         =   "Arquivo para Fusão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FusaoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFusao_Click()
Dim sFile As String
    With DlgFusao
        .InitDir = g_Parametros.DiretorioTransmissao
        .DialogTitle = "Localizar Arquivo de Fusão"
        .CancelError = False
        .Filter = "Arquivo de Fusão (*.GCC)|*.GCC"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    TxtFusao.Text = sFile
    TxtFusao.SetFocus
End Sub

Private Sub cmdOK_Click()
 FusaoDialog.Visible = False
 Call recepcao.FusaoAutomatica
End Sub

