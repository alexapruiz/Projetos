Attribute VB_Name = "ProUtil"
Option Explicit

Declare Function UT_FILTRABMP Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String, ByVal LimiteCorte As Long, ByVal ImageType As Long)
' ImageType
' 1 - bmp
' 2 - TIFFG4
Declare Function UT_Trim Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String) As Integer

Declare Function UT_CarimbaBMP Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String, ByVal StringCarimbo As String, ByVal tTop As Long, ByVal tLeft As Long, ByVal negrito As Long, ByVal Transparente As Long, ByVal angulo As Long) As Long

Declare Function UT_TIFF2BMP Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String) As Long

Declare Function UT_BMP2TIFF Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String) As Long

Declare Function UT_JPG2BMP Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String) As Long

Declare Function UT_CortaBMP Lib "PROUT_32.DLL" (ByVal FileOrigem As String, ByVal FileDestino As String, ByVal larg As Long, ByVal alt As Long, ByVal topo As Long, ByVal rod As Long) As Long

Declare Function UT_DesmembraTiff Lib "PROUT_32.DLL" (ByVal PathImagem As String, ByVal ArqTxt As String, ByVal num_lote As Long, ByVal ImgInicialCanon As Long, ByVal Estacao As Long) As Long

Declare Function UT_GaugeCanonInit Lib "PROUT_32.DLL" (ByVal Top As Long, ByVal Left As Long) As Long

Declare Function UT_GaugeCanon Lib "PROUT_32.DLL" () As Long

Declare Sub UT_DestroyGaugeCanon Lib "PROUT_32.DLL" ()

