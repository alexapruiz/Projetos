VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Funcoes da dll VIPSDLL.Dll do Unibanco

' Inicializa as Dlls
Private Declare Function VIPS_Init Lib "VIPSDLL.DLL" () As Long
' Libera as Dlls
Private Declare Sub VIPS_Done Lib "VIPSDLL.DLL" ()
' Seleciona a porta serial (Default = 1)
Private Declare Sub VIPS_SetComPort Lib "VIPSDLL.DLL" (ByVal ComPort As Long)
' Seleciona a resolucao ( 100 ou 200 DPI ) (Default = 100)
Private Declare Sub VIPS_SetResolution Lib "VIPSDLL.DLL" (ByVal Res As Long)
' Seleciona leitora ( 1 = CMC7; 2 = Cod. Barras; 3 = Ambas ) (Default = 3)
Private Declare Sub VIPS_SetReader Lib "VIPSDLL.DLL" (ByVal Reader As Long)
' Seleciona a quantidade de Escaninhos ( Nro. de escaninhos deve ser impar )
' (Default = 1)
Private Declare Sub VIPS_SetBoxes Lib "VIPSDLL.DLL" (ByVal Boxes As Long)
' Seleciona a qtde maxima de docs por escanhinho (Default = 120)
Private Declare Sub VIPS_SetMaxDocBox Lib "VIPSDLL.DLL" (ByVal MaxDocBox As Long)
' Seleciona o escanhinho default ( Default = 0 : Nenhum )
Private Declare Sub VIPS_SetBoxDefault Lib "VIPSDLL.DLL" (ByVal BoxDefault As Long)
' Seleciona o tipo de imagem gerada ( 1 = BMP; 3 = JPG ) (Default = 3)
Private Declare Sub VIPS_SetImageType Lib "VIPSDLL.DLL" (ByVal ImageType As Long)
' Seleciona o fator de compressao do JPG ( Default = 30 )
Private Declare Sub VIPS_SetCompress Lib "VIPSDLL.DLL" (ByVal Fator As Long)
' Seleciona se deve cortar as bordas
' ( Valor-> 0 a 255, Zero nao corta. Quanto mais alto, mais corta )
' (Default = 75)
Private Declare Sub VIPS_SetCutBords Lib "VIPSDLL.DLL" (ByVal Valor As Long)
' Seleciona o diretorio onde serao gravadas as imagens
Private Declare Sub VIPS_SetImageDirectory Lib "VIPSDLL.DLL" (ByVal diretorio As String)
' Seleciona o arq de conf. da camera ( nao precisa do caminho )
Private Declare Sub VIPS_SetCameraFile Lib "VIPSDLL.DLL" (ByVal CameraFile As String)
' Executa a captura
Private Declare Function VIPS_Captura Lib "VIPSDLL.DLL" (ByVal AgProc As Long, ByVal Lote As Long, ByVal ArqRetorno As String, ByVal Append As Long) As Long
' Reseta as Dlls
Private Declare Function VIPS_Reset Lib "VIPSDLL.DLL" () As Long


Public Function Init() As Long
    Init = VIPS_Init()
End Function

Public Sub Done()
    VIPS_Done
End Sub

Public Sub SetComPort(ByVal ComPort As Long)
    VIPS_SetComPort (ComPort)
End Sub

Public Sub SetResolution(ByVal Res As Long)
    VIPS_SetResolution (Res)
End Sub

Public Sub SetBoxes(ByVal Boxes As Long)
    VIPS_SetBoxes (Boxes)
End Sub

Public Sub SetMaxDocBox(ByVal MaxDocBox As Long)
    VIPS_SetMaxDocBox (MaxDocBox)
End Sub

Public Sub SetBoxDefault(ByVal BoxDefault As Long)
    VIPS_SetBoxDefault (BoxDefault)
End Sub

Public Sub SetImageType(ByVal ImageType As Long)
    VIPS_SetImageType (ImageType)
End Sub

Public Sub SetCutBords(ByVal Valor As Long)
    VIPS_SetCutBords (Valor)
End Sub

Public Sub SetImageDirectory(ByVal diretorio As String)
    VIPS_SetImageDirectory (diretorio)
End Sub

Public Sub SetCameraFile(ByVal CameraFile As String)
    VIPS_SetCameraFile (CameraFile)
End Sub

Public Function Captura(ByVal AgProc As Long, ByVal Lote As Long, ByVal ArqRetorno As String, ByVal Append As Long) As Long
    Captura = VIPS_Captura(AgProc, Lote, ArqRetorno, Append)
End Function

Public Function Reset()
    Reset = VIPS_Reset()
End Function

Public Sub SetCompress(ByVal Fator As Long)
    VIPS_SetCompress (Fator)
End Sub
