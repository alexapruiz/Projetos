Attribute VB_Name = "ScannerInterface"
Option Explicit

'*******************************************************************
'***       Declaraçoes das Funções de ScannerInterface.DLL       ***
'*******************************************************************

Declare Function SC_SetScanner Lib "ScannerInterface.DLL" (ByVal lngScanner As Long) As Long
Declare Function SC_Init Lib "ScannerInterface.DLL" () As Long
Declare Function SC_DeInit Lib "ScannerInterface.DLL" () As Long
Declare Function SC_Eject Lib "ScannerInterface.DLL" () As Long
Declare Function SC_SendPocket Lib "ScannerInterface.DLL" (ByVal lngEscaninho As Long) As Long
Declare Function SC_SetPocketDefault Lib "ScannerInterface.DLL" (ByVal lngPocket As Long) As Long
Declare Function SC_SetNumPockets Lib "ScannerInterface.DLL" (ByVal lngNumPockets As Long) As Long
Declare Function SC_SetComPort Lib "ScannerInterface.DLL" (ByVal lngComPort As Long) As Long
Declare Function SC_SetImageType Lib "ScannerInterface.DLL" (ByVal lngImageType As Long) As Long
Declare Function SC_SetImageDPI Lib "ScannerInterface.DLL" (ByVal lngDPI As Long) As Long
Declare Function SC_SetDocWidth Lib "ScannerInterface.DLL" (ByVal lngLargura As Long) As Long
Declare Function SC_SetDocHeight Lib "ScannerInterface.DLL" (ByVal lngAltura As Long) As Long
Declare Function SC_SetCodeLine Lib "ScannerInterface.DLL" (ByVal lngCodeLine As Long) As Long
Declare Function SC_SetNumPages Lib "ScannerInterface.DLL" (ByVal lngNumPages As Long) As Long
Declare Function SC_SetAppend Lib "ScannerInterface.DLL" (ByVal lngFlagAppend As Long) As Long
Declare Function SC_DoubleDetect Lib "ScannerInterface.DLL" (ByVal lngPrecisao As Long) As Long
Declare Function SC_SetImageQuality Lib "ScannerInterface.DLL" (ByVal lngImageQuality As Long) As Long
Declare Function SC_SetGaugePos Lib "ScannerInterface.DLL" (ByVal lngTop As Long, ByVal lngLeft As Long) As Long
Declare Function SC_SetImagePath Lib "ScannerInterface.DLL" (ByVal strImagePath As String) As Long
Declare Function SC_GetErrorMessage Lib "ScannerInterface.DLL" (ByVal lngErrorCode As Long, ByVal strErrorMessage As String) As Long
Declare Function SC_AcquireSingle Lib "ScannerInterface.DLL" (ByVal strFileFront As String, ByVal strFileBack As String, ByVal strCodeLine As String) As Long
'Declare Function SC_AcquireBatch Lib "ScannerInterface.DLL" (ByVal lngNumInicial As Long, ByVal strFileRetorno As String, ByVal lngEstacao As Long) As Long
Declare Function SC_AcquireBatch Lib "ScannerInterface.DLL" (ByVal lngAgencia As Long, ByVal lngLote As Long, ByVal lngSeqInicio As Long, ByVal strFileRetorno As String, ByVal lngEstacao As Long) As Long

'******************************************************
'***                    Constantes                  ***
'******************************************************


'Scanners
Global Const MC93 = 1
Global Const LS500 = 2
Global Const VISIONSHAPE = 3
Global Const BUIC1500 = 4
Global Const SB1600 = 5
Global Const RDS3000 = 6

'Pages
Global Const FRONT = 1
Global Const BACK = 2
Global Const FRONTBACK = 3

'ImageType
Global Const IMAGE_BMP = 1
Global Const IMAGE_TIF = 2
Global Const IMAGE_JPG = 3
Global Const IMAGE_JPGCOLOR = 4

'DPI
Global Const DPI_100 = 100
Global Const DPI_200 = 200

'CodeLine
Global Const CMC7 = 1
Global Const BARCODE = 2
Global Const ALL_CODES = 3

'Retorno de Método
Global Const SC_OK = 1
Global Const SC_Erro = 0

'Variaveis de configuração da Nova DLL (Vips)
Private Type tpParamVips
    NumBoxes As Long
    MaxDocBox As Long
    BoxDefault As Long
End Type
Public tSC_ParamDLL As tpParamVips
'


Public Sub ScanMessageErr(iRetDoc As Long)
        
Dim iRet As Long
Dim strErro As String * 128
    strErro = Space(64)
    iRet = SC_GetErrorMessage(iRetDoc, strErro)
    MsgBox strErro, vbOKOnly, "Erro: " & str(iRetDoc)

End Sub

Public Function InicializarVips() As Boolean
    
Dim iRet As Long

InicializarVips = False

    If bInicializou = False Then
        iRet = SC_SetScanner(MC93)
        If iRet = 1 Then
            iRet = SC_SetNumPockets(tSC_ParamDLL.NumBoxes)
            If iRet = 1 Then
                iRet = SC_SetPocketDefault(tSC_ParamDLL.BoxDefault)
                If iRet = 1 Then
                    iRet = SC_SetComPort(1)
                    If iRet = 1 Then
                        iRet = SC_SetImageType(IMAGE_JPG)
                        If iRet = 1 Then
                            iRet = SC_SetImageDPI(DPI_100)
                            If iRet = 1 Then
                                iRet = SC_SetDocWidth(810)
                                If iRet = 1 Then
                                    iRet = SC_SetDocHeight(420)
                                    If iRet = 1 Then
                                        iRet = SC_SetCodeLine(ALL_CODES)
                                        If iRet = 1 Then
                                            iRet = SC_SetNumPages(FRONTBACK)
                                            If iRet = 1 Then
                                                iRet = SC_DoubleDetect(2)
                                                If iRet = 1 Then
                                                    iRet = SC_SetImageQuality(50)
                                                    If iRet = 1 Then
                                                        iRet = SC_SetImagePath(Geral.DiretorioImagens)
                                                    End If  'SC_SetImageQuality
                                                End If  'SC_DoubleDetect
                                            End If  'SC_SetNumPages
                                        End If  'SC_SetCodeLine
                                    End If  'SC_SetDocHeight
                                End If  'SC_SetDocWidth
                            End If  'SC_SetImageDPI
                        End If  'SC_SetImageType
                    End If  'SC_SetComPort
                End If  'SC_SetPocketDefault
            End If  'SC_SetNumPockets
        End If  'SC_SetScanner
        
        If iRet <> 1 Then
            Call ScanMessageErr(iRet)
            Exit Function
        End If
        iRet = SC_Init()
        If iRet <> 1 Then
            Call ScanMessageErr(iRet)
            Exit Function
        Else
            bInicializou = True
        End If
    End If  'bInicializou = False

InicializarVips = True

End Function

