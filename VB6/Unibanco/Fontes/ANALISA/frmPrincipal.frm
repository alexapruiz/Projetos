VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   " Analisador de Utilização de Scanner"
   ClientHeight    =   2556
   ClientLeft      =   132
   ClientTop       =   612
   ClientWidth     =   3744
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnu_Abrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnu_Hifen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuASair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuProcessamento 
      Caption         =   "&Processamento"
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "&Sobre"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub Form_Load()
   Dim nInd As Integer
   
   GetComputerName sMaquina, 30
 
   Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor

   For nInd = 1 To Len(sMaquina)
      If Asc(Mid(sMaquina, nInd, 1)) = 0 Then
         Mid(sMaquina, nInd, 1) = " "
      End If
   Next nInd
   
'    Printer.Orientation = 1
'
'    frmPrint.Width = Printer.ScaleWidth
'    frmPrint.Height = Printer.ScaleHeight
'    frmPrint.Show vbModal

'    frmImprimeDatas.cmdImprimir_Click

End Sub

Private Sub mnu_Abrir_Click()
    frmImprimeDatas.Show vbModal
End Sub


Private Sub mnuASair_Click()
   End
End Sub
Private Sub mnuGPEstacao_Click()
   On Error Resume Next
   
   'frmGraficoProdutividade.Show vbModal, Me
End Sub
Private Sub mnuProcessamento_Click()
   frmProcessamento.Show vbModal, Me
End Sub
Private Sub mnuSobre_Click()
   frmSobre.Show vbModal, Me
End Sub
