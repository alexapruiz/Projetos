Attribute VB_Name = "KOFAX_SCAN"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Copyright (c) 1993-1996, Kofax Image Products. All rights reserved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'       File: kscan_m.bas
'    Purpose: Visual Basic Constants for use with KSCAN.VBX custom control
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' tab=4 '

Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Control Specific Properties
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Unit Settings
Global Const KSUNITENGLISH = 0
Global Const KSUNITMETRIC = 1

' Action Settings
Global Const KSACTIONNONE = 0
Global Const KSACTIONRESERVE = 1
Global Const KSACTIONUNRESERVE = 2
Global Const KSACTIONRESET = 3
Global Const KSACTIONSTART = 4
Global Const KSACTIONSTARTASYNC = 5
Global Const KSACTIONSTOP = 6
Global Const KSACTIONOPENDIALOG = 7
Global Const KSACTIONREADINI = 8
Global Const KSACTIONWRITEINI = 9
Global Const KSACTIONSETSETTINGS = 10
Global Const KSACTIONGETSETTINGS = 11
Global Const KSACTIONSENDCUSTOMCMD = 12
Global Const KSACTIONSTOPIMMEDIATE = 13

' ActionResetMode Settings
Global Const KSACTIONRESETMODENOWAIT = 0
Global Const KSACTIONRESETMODEWAITFOROK = 1
Global Const KSACTIONRESETMODETIMEDWAIT = 2
Global Const KSACTIONRESETMODENODLG = 3

' ActiveDialog Settings
Global Const KSDIALOGABOUT = 0
Global Const KSDIALOGSELECTDEVICE = 1
Global Const KSDIALOGDEVICESETTINGS = 2
Global Const KSDIALOGSTORAGE = 3
Global Const KSDIALOGADVANCED = 4
Global Const KSDIALOGDEVICEOPTIONS = 5

' ScanColorMode Settings
Global Const KSSCANCOLORMODEBITONAL = &H1
Global Const KSSCANCOLORMODE16MCOLOR = &H2
Global Const KSSCANCOLORMODE256COLOR = &H4
Global Const KSSCANCOLORMODE256GRAY = &H8
Global Const KSSCANCOLORMODE16GRAY = &H10
Global Const KSSCANCOLORMODEALLCAP = &H1F

' ActiveDevice Settings
Global Const KSACTIVEDEVICESCANNER = 0
Global Const KSACTIVEDEVICEDISK = 1

' DeviceMethod Settings
Global Const KSDEVICEMETHODSINGLE = 0
Global Const KSDEVICEMETHODBATCH = 1

' DeviceStatus Settings
Global Const KSDEVICESTATUSOUTOFPAPER = &H1
Global Const KSDEVICESTATUSPAPERJAM = &H2

' ScanDestination Settings
Global Const KSDESTBIN1 = &H1
Global Const KSDESTBIN2 = &H2
Global Const KSDESTALLCAP = &H3

' ScanDirection Settings
Global Const KSSCANDIRECTPORTRAIT = &H1
Global Const KSSCANDIRECTLANDSCAPE = &H2
Global Const KSSCANDIRECTPANEL = &H4
Global Const KSSCANDIRECTALLCAP = &H7

' ScanMode Settings
Global Const KSSCANMODEPANEL = &H1
Global Const KSSCANMODELINE = &H2
Global Const KSSCANMODEPHOTO = &H4
Global Const KSSCANMODEMIXED = &H8
Global Const KSSCANMODEALLCAP = &HF

' ScanSideMode Settings
Global Const KSSCANSIDEMODEFRONT = 0
Global Const KSSCANSIDEMODEBACK = 1

' ScanSource Settings
Global Const KSSOURCEFLATBED = &H1
Global Const KSSOURCEADF = &H2
Global Const KSSOURCEPANEL = &H4
Global Const KSSOURCEMANUAL = &H8
Global Const KSSOURCEALLCAP = &HF

' DeviceCache Settings
Global Const KSCACHENONE = 0
Global Const KSCACHE1 = 1
Global Const KSCACHE2 = 2
Global Const KSCACHE3 = 3
Global Const KSCACHE4 = 4
Global Const KSCACHE5 = 5
Global Const KSCACHE6 = 6

' PEPedalStatus Settings
Global Const KSFOOT0 = &H100
Global Const KSFOOT1 = &H200
Global Const KSFOOT2 = &H400
Global Const KSFOOTALLCAP = &H700

' FujitsuIPCDTCMode Settings
Global Const KSFUJITSUIPCDTCMODENORMAL = 0
Global Const KSFUJITSUIPCDTCMODESIMPLIFIED = 1
Global Const KSFUJITSUIPCDTCMODESTANDARD = 2

' FujitsuIPCEmphasis Settings
Global Const KSFUJITSUIPCEMPHASISNONE = 0
Global Const KSFUJITSUIPCEMPHASISLARGE = 1
Global Const KSFUJITSUIPCEMPHASISMIDDLE = 2
Global Const KSFUJITSUIPCEMPHASISSMALL = 3
Global Const KSFUJITSUIPCEMPHASISSMOOTHING = 4

' FujitsuIPCVariance Settings
Global Const KSFUJITSUIPCVARIANCE0 = 0
Global Const KSFUJITSUIPCVARIANCE1 = 1
Global Const KSFUJITSUIPCVARIANCE2 = 2
Global Const KSFUJITSUIPCVARIANCE3 = 3
Global Const KSFUJITSUIPCVARIANCE4 = 4
Global Const KSFUJITSUIPCVARIANCE5 = 5
Global Const KSFUJITSUIPCVARIANCE6 = 6
Global Const KSFUJITSUIPCVARIANCE7 = 7

' FujitsuIPCPreset Settings
Global Const KSFUJITSUIPCPRESETCUSTOM = 0
Global Const KSFUJITSUIPCPRESET1 = 1
Global Const KSFUJITSUIPCPRESET2 = 2
Global Const KSFUJITSUIPCPRESET3 = 3
Global Const KSFUJITSUIPCPRESET4 = 4
Global Const KSFUJITSUIPCPRESET5 = 5
Global Const KSFUJITSUIPCPRESET6 = 6

' FujitsuIPCGamma Settings
Global Const KSFUJITSUIPCGAMMANORMAL = 0
Global Const KSFUJITSUIPCGAMMASOFT = 1
Global Const KSFUJITSUIPCGAMMASHARP = 2

' FujitsuIPCDTCCurve Settings
Global Const KSFUJITSUIPCDTCCURVE0 = 0
Global Const KSFUJITSUIPCDTCCURVE1 = 1
Global Const KSFUJITSUIPCDTCCURVE2 = 2
Global Const KSFUJITSUIPCDTCCURVE3 = 3
Global Const KSFUJITSUIPCDTCCURVE4 = 4
Global Const KSFUJITSUIPCDTCCURVE5 = 5
Global Const KSFUJITSUIPCDTCCURVE6 = 6
Global Const KSFUJITSUIPCDTCCURVE7 = 7

' RicohIPCap Settings
Global Const KSRICOHIPGAMMA0 = 0
Global Const KSRICOHIPGAMMA1 = 1
Global Const KSRICOHIPGAMMA2 = 2
Global Const KSRICOHIPGAMMA3 = 3
Global Const KSRICOHIPGAMMA4 = 4
Global Const KSRICOHIPGAMMA5 = 5
Global Const KSRICOHIPGAMMA6 = 6
Global Const KSRICOHIPGAMMA7 = 7
Global Const KSRICOHIPGAMMA8 = 8
Global Const KSRICOHIPGAMMA9 = 9
Global Const KSRICOHIPGAMMA10 = 10
Global Const KSRICOHIPGAMMA31 = 31

' RicohIPGamma Settings
Global Const KSRICOHIPBACKTRACK0 = 0
Global Const KSRICOHIPBACKTRACK1 = 1
Global Const KSRICOHIPBACKTRACK2 = 2
Global Const KSRICOHIPBACKTRACK3 = 3
Global Const KSRICOHIPBACKTRACK4 = 4
Global Const KSRICOHIPBACKTRACK5 = 5
Global Const KSRICOHIPBACKTRACK6 = 6
Global Const KSRICOHIPBACKTRACK7 = 7
Global Const KSRICOHIPBACKTRACK12 = 12
Global Const KSRICOHIPBACKTRACK13 = 13
Global Const KSRICOHIPBACKTRACK14 = 14
Global Const KSRICOHIPBACKTRACK15 = 15

' RicohIPMatrixSelect Settings
Global Const KSRICOHIPMATRIXSELECT0 = 0
Global Const KSRICOHIPMATRIXSELECT1 = 1
Global Const KSRICOHIPMATRIXSELECT2 = 2
Global Const KSRICOHIPMATRIXSELECT3 = 3
Global Const KSRICOHIPMATRIXSELECT4 = 4
Global Const KSRICOHIPMATRIXSELECT5 = 5
Global Const KSRICOHIPMATRIXSELECT6 = 6
Global Const KSRICOHIPMATRIXSELECT7 = 7

' RicohIPSPPixel Settings
Global Const KSRICOHIPSPPIXEL0 = 0
Global Const KSRICOHIPSPPIXEL1 = 1
Global Const KSRICOHIPSPPIXEL2 = 2
Global Const KSRICOHIPSPPIXEL3 = 3

' RicohIPChangeRateSense Settings
Global Const KSRICOHIPCHANGERATESENSE0 = 0
Global Const KSRICOHIPCHANGERATESENSE1 = 1
Global Const KSRICOHIPCHANGERATESENSE2 = 2
Global Const KSRICOHIPCHANGERATESENSE3 = 3

' RicohIPChangeRateAlgo Settings
Global Const KSRICOHIPCHANGERATEALGO0 = 0
Global Const KSRICOHIPCHANGERATEALGO1 = 1

' BHACEFunction Settings
Global Const KSBHACEFUNCTIONLIGHTTEXT = 0
Global Const KSBHACEFUNCTIONPHOTO = 1
Global Const KSBHACEFUNCTIONLINEART = 2
Global Const KSBHACEFUNCTIONMIXED = 3
Global Const KSBHACEFUNCTIONDARKTEXT = 4

' BHACESensitivity Settings
Global Const KSBHACESENSITIVITY0 = 0
Global Const KSBHACESENSITIVITY1 = 1
Global Const KSBHACESENSITIVITY2 = 2
Global Const KSBHACESENSITIVITY3 = 3
Global Const KSBHACESENSITIVITY4 = 4
Global Const KSBHACESENSITIVITY5 = 5
Global Const KSBHACESENSITIVITY6 = 6
Global Const KSBHACESENSITIVITY7 = 7



' declarações do KUTIL.BAS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Copyright (c) 1993-1996 Kofax Image Products.  All rights reserved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'       File: kfutil.bas
'    Purpose: Kofax utility functions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' tab=4



'***** Global variable definitions. (All globals start with g_)
'Global g_ActiveKView As KView         ' Currently active KView
Global g_bLeftSide As Boolean         ' Side for Analysis window

Global g_PrevTime As Single           ' Clock count for previous scan
Global g_ElapTime As Single           ' Running elapsed time for batch scanning
Global g_StartTime As Single          ' Start time to calculate PPM
Global g_PEPageCount As Single        ' Page end page counter
Global g_PSPageCount As Single        ' Page start page counter
Global g_TotalPages As Single         ' Total number of pages since starting demo
Global g_bOutOfPaper As Boolean       ' Out of paper message on status line

'***** Page end file parameters (to disk)
Global g_PEFileNamePrefix As String * 3         ' Prefix for scanned file names
Global g_PEFileNameExtension As String * 3      ' Extension for scanned file names
Global g_PEFilePath As String                   ' Path to store scanned images into
Global g_PEWriteMode As Integer                 ' KGOverwrite, KGAppend, KGInsert
Global g_PEPage As Integer                      ' Page number for overwrite and insert

'***** Page start file parameters (from disk)
Global g_PSFileArray() As String            ' Up to 256 files to be read from disk
Global g_PSFileNamePrefix As String * 3     ' Prefix for scanned file names
Global g_PSFileNameExtension As String * 3  ' Extension for scanned file names
Global g_PSFilePath As String               ' Path to store scanned images into
Global g_IniFileName As String              ' Path for INI file, .exe runtime directory
Global g_Endorse As Boolean                 ' TRUE if Endorse is Capable and Enabled
Global g_PSEndorseText As String            ' Text for scanner endorse
Global g_nMustRestoreAnalysis As Integer    ' Number of times analysis has been hidden

'***** Print parameters used to support multipage printing
Global g_bPrintAll As Boolean               ' TRUE if print all pages is selected
Global g_nPrintPageCount As Integer         ' Number of pages to print
Global g_nPrintCurrentPage As Integer       ' Page to print in print pagestart event

'***** Analysis log file variables
Global g_strLogFileName As String           ' Filename to log to

'***** Application title
#If Win32 Then
Global Const gc_Title = "DEMO - CODIGO DE BARRAS 32 BITS"
#ElseIf Win16 Then
Global Const gc_Title = "DEMO - CODIGO DE BARRAS 16 BITS"
#Else
    '***** FORCE ERROR! This version of VBDEMO only works in 32-bit mode.
    Error (UNSUPPORTEDOS)
#End If

Global Const LF_FACESIZE = 32

Type LOGFONT
    lfHeight As Integer
    lfWidth As Integer
    lfEscapement As Integer
    lfOrientation As Integer
    lfWeight As Integer
    lfItalic As String * 1
    lfUnderline As String * 1
    lfStrikeOut As String * 1
    lfCharSet As String * 1
    lfOutPrecision As String * 1
    lfClipPrecision As String * 1
    lfQuality As String * 1
    lfPitchAndFamily As String * 1
    lfFaceName As String * LF_FACESIZE
End Type

'*****************************************************************************
' WINDOWS API
'*****************************************************************************
#If Win32 Then
        Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
        Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
        Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
        Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
        Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#ElseIf Win16 Then
        Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
        Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
        Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
        Declare Function IsWindowVisible Lib "User" (ByVal hWnd As Integer) As Integer
        Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer
#Else
    '***** FORCE ERROR! Unsupported operating system.
    Error (UNSUPPORTEDOS)
#End If

'*****************************************************************************
' IMPORTED FUNCTION API
'*****************************************************************************


'*****************************************************************************
' Global variables
'*****************************************************************************
Global Const SUCCESS = 0
Global Const INVALID_SETTING = 1        ' Invalid setting for property
Global Const INVALID_FILE_FORMAT = 321  ' Standard VB error
Global Const INVALID_FILE_NAME = 52     ' Standard VB error Bad file name or number

