Attribute VB_Name = "LS500_DLL"
Option Explicit

'---- VARIÁVEIS USADAS PARA LS500 ----
Global NumEqs As Integer

Global String1 As String * 25
Global String2 As String * 25
Global String3 As String * 25


'--- API ---
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
'Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal w_hwnd As Integer) As Integer






