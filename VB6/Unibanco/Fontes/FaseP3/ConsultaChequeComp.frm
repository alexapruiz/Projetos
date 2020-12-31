VERSION 5.00
Begin VB.Form ConsultaChequeComp 
   Caption         =   "Consulta de Cheques Compensação"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ConsultaChequeComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()

    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(28)
    
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

End Sub
