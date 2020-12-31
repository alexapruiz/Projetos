Attribute VB_Name = "LS500_DLL"
Option Explicit

'*******************************************************
'***            INTERFACE LS-500 - 32 Bits           ***
'*******************************************************

Declare Function LS_Init Lib "LS500_32.DLL" () As Long

Declare Function LS_Deinit Lib "LS500_32.DLL" () As Long

Declare Function LS_Reset Lib "LS500_32.DLL" () As Long

Declare Function LS_ProcuraLS500 Lib "LS500_32.DLL" (ByVal String1 As String, ByVal String2 As String, ByVal String3 As String) As Long

Declare Function LS_LS500Atual Lib "LS500_32.DLL" (ByVal Equip As Long) As Long

Declare Function LS_GetID Lib "LS500_32.DLL" (ByVal Equip As Long) As Long

Declare Function LS_GetLUN Lib "LS500_32.DLL" (ByVal Equip As Long) As Long

Declare Function LS_Digitaliza Lib "LS500_32.DLL" (ByVal FileFrente As String, ByVal FileVerso As String, ByVal StringCMC7 As String, ByVal FlagSepara As Long) As Long

Declare Function LS_SetGaugePos Lib "LS500_32.DLL" (ByVal xTop As Long, ByVal xLeft As Long) As Long

Declare Function LS_SetNumGauges Lib "LS500_32.DLL" (ByVal nGauges As Long) As Long

Declare Function LS_Digitaliza1 Lib "LS500_32.DLL" (ByVal DirIMG As String, ByVal NumInicioIMG As Long, ByVal Estacao As Long) As Long

Declare Function LS_Digitaliza2 Lib "LS500_32.DLL" (ByVal DirIMG As String, ByVal NumInicio1IMG As Long, ByVal NumInicio2IMG As Long, ByVal Estacao As Long) As Long

Declare Function LS_Digitaliza3 Lib "LS500_32.DLL" (ByVal DirIMG As String, ByVal NumInicio1IMG As Long, ByVal NumInicio2IMG As Long, ByVal NumInicio3IMG As Long, ByVal Estacao As Long) As Long

Declare Function LS_Lapso Lib "LS500_32.DLL" (ByVal nTics As Long) As Long

Declare Function LS_SetAltura Lib "LS500_32.DLL" (ByVal Altura As Long) As Long

Declare Function LS_GetRecursos Lib "LS500_32.DLL" (ByVal StrRecursos As String) As Long

Declare Function LS_SetImage Lib "LS500_32.DLL" (ByVal f_v As Long) As Long

Declare Function LS_SetLeitora Lib "LS500_32.DLL" (ByVal CMC7ouBarCode As Long) As Long

Declare Function LS_SetCarimbo Lib "LS500_32.DLL" (ByVal f_v As Long) As Long

Declare Function LS_SetString Lib "LS500_32.DLL" (ByVal StrEnvia As String, ByVal Bold As Long) As Long

Declare Function LS_SetDocRigido Lib "LS500_32.DLL" (ByVal T_F As Long) As Long

Declare Function LS_SetSepara Lib "LS500_32.DLL" (ByVal T_F As Long) As Long

Declare Function LS_SetImageColor Lib "LS500_32.DLL" (ByVal Color As Long) As Long

Declare Function LS_SetImageDPI Lib "LS500_32.DLL" (ByVal DPI As Long) As Long

Declare Function LS_SetCutLimit Lib "LS500_32.DLL" (ByVal Limite As Long) As Long

Declare Function LS_SetLimitAlternate Lib "LS500_32.DLL" (ByVal Limite1 As Long, ByVal Limite2 As Long, ByVal Limite3 As Long, ByVal Limite4 As Long, ByVal Limite5 As Long) As Long

Declare Function LS_SetBankAlternate Lib "LS500_32.DLL" (ByVal Banco1 As Long, ByVal Banco2 As Long, ByVal Banco3 As Long, ByVal Banco4 As Long, ByVal Banco5 As Long) As Long

Declare Function LS_SetTimeOut Lib "LS500_32.DLL" (ByVal MiliSeg As Long) As Long

Declare Function LS_SetMaxErrorCMC7 Lib "LS500_32.DLL" (ByVal MaxErr As Long) As Long

Declare Function LS_SetFileName Lib "LS500_32.DLL" (ByVal FileName As String) As Long

Declare Function LS_SetAppend Lib "LS500_32.DLL" (ByVal T_F As Long) As Long

Declare Function LS_GetVersao Lib "LS500_32.DLL" () As Double

' Funções para LS + Canon

Declare Function LS_SetCanon Lib "LS500_32.DLL" (ByVal T_F As Long) As Long

Declare Function LS_GetFileRetorno Lib "LS500_32.DLL" (ByVal FileName As String) As Long

Declare Function LS_GetLastImage Lib "LS500_32.DLL" () As Long

Declare Function LS_GetNumLote Lib "LS500_32.DLL" () As Long


'---- VARIÁVEIS USADAS PARA LS500 ----
Global NumEqs As Integer

Global String1 As String * 25
Global String2 As String * 25
Global String3 As String * 25


'--- API ---
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
'Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal w_hwnd As Integer) As Integer






