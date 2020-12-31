VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4956
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6528
   LinkTopic       =   "Form1"
   ScaleHeight     =   4956
   ScaleWidth      =   6528
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim ObjCalc As Object
    Dim ObjCalc2 As Object
    
    Set ObjCalc2 = CreateObject("FuncoesApoio.CalculoDigito")
    Set ObjCalc = CreateObject("FuncoesApoio.Criptografia")
    
    'MsgBox ObjCalc2.modulo10(1)
    'MsgBox ObjCalc2.modulo11(2)
    'MsgBox ObjCalc.encripta("a")
    
    Set ObjCalc = Nothing
    Set ObjCalc2 = Nothing
End Sub
