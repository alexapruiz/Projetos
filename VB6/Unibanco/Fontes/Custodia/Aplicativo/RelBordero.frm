VERSION 5.00
Begin VB.Form RelBordero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Borderôs"
   ClientHeight    =   2175
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   732
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4092
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1332
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   372
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4092
      Begin VB.TextBox txtBordero 
         Height          =   288
         Left            =   1920
         MaxLength       =   19
         TabIndex        =   3
         Top             =   600
         Width           =   1932
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Borderô :"
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1572
      End
   End
End
Attribute VB_Name = "RelBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_Change()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdImprimir_Click()

Dim rsPesquisaIDBordero            As New ADODB.Recordset

Dim vIdBordero                     As Double

Dim SemRegistros                   As Integer

Dim Selecao                        As New Custodia.Selecionar

Screen.MousePointer = vbHourglass

txtBordero.Text = FormataString(txtBordero.Text, "0", txtBordero.MaxLength, True)

Set rsPesquisaIDBordero = g_cMainConnection.Execute(Selecao.GetNumeroIdBordero(Geral.DataProcessamento, txtBordero.Text))

If Not rsPesquisaIDBordero.EOF Then

    vIdBordero = rsPesquisaIDBordero!IdBordero
    
    Principal.CrystalReport.CopiesToPrinter = 1
    Principal.CrystalReport.ReportFileName = App.path & "\Reports\RelGerenciamentoCheques.rpt"
    Principal.CrystalReport.SelectionFormula = "{Bordero.IdBordero} = " + Trim(Str(vIdBordero))
    Principal.CrystalReport.WindowState = crptMaximized
    Principal.CrystalReport.WindowTitle = "Emissão do Relatório de Borderôs"
    Principal.CrystalReport.Action = 0

Else
    SemRegistros = MsgBox("Borderô Inválido", vbExclamation, "Relatório de Borderôs Para Transmissão")
End If

Screen.MousePointer = Default

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdSair_Click()

Unload Me

End Sub

Private Sub txtBordero_GotFocus()
    SelecionarTexto txtBordero
End Sub

Private Sub txtBordero_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       Call cmdImprimir_Click
    Else
       SoNumero KeyAscii
    End If
    
End Sub

Private Sub txtBordero_LostFocus()

If Trim(txtBordero.Text) = "" Then Exit Sub

txtBordero.Text = FormataString(txtBordero.Text, "0", txtBordero.MaxLength, True)
    
End Sub

Private Sub txtBordero_Validate(Cancel As Boolean)

    Dim stxtNumeroBordero   As String * 19
    Dim sCPFCGC             As String * 12
    Dim sNumBordero         As String * 6
    Dim sDV                 As String * 1
    
    On Error GoTo NumeroInvalido
    
    If Trim(txtBordero.Text) = "" Then Exit Sub
    
    stxtNumeroBordero = FormataString(txtBordero.Text, "0", txtBordero.MaxLength, True)
    
    sCPFCGC = Left(stxtNumeroBordero, 12)
    
    sNumBordero = Mid(stxtNumeroBordero, 13, 6)
    
    sDV = Right(stxtNumeroBordero, 1)
    
    If Val(sNumBordero) = 0 Then GoTo NumeroInvalido
    
    If Not Modulo11Simplificado(sNumBordero & sDV) Then GoTo NumeroInvalido

    Exit Sub
NumeroInvalido:

    MsgBox "Número de Borderô inválido.", vbExclamation, Me.Caption
    txtBordero.Text = ""
    Cancel = True
    SelecionarTexto txtBordero
    Exit Sub

End Sub
