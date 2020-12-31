Attribute VB_Name = "VIPSDLL"
Option Explicit

' Funcoes da Dll VIPS_DLL.Dll da Proservvi
Declare Function MC93_Init Lib "VIPS_DLL.DLL" () As Long
Declare Function MC93_DeInit Lib "VIPS_DLL.DLL" () As Long
Declare Function MC93_SetComPort Lib "VIPS_DLL.DLL" (ByVal CommPort As Long) As Long ' 1 ou 2
Declare Function MC93_SetDPI Lib "VIPS_DLL.DLL" (ByVal DPI As Long) As Long      ' 100 ou 200
Declare Function MC93_SetLeitora Lib "VIPS_DLL.DLL" (ByVal Leitora As Long) As Long ' 1=CMC-7 2=BarCode 3=Ambas
Declare Function MC93_SetAltura Lib "VIPS_DLL.DLL" (ByVal Altura As Long) As Long ' 310 0u 420
Declare Function MC93_SetAppend Lib "VIPS_DLL.DLL" (ByVal Apnd As Long) As Long
Declare Function MC93_SetImagem Lib "VIPS_DLL.DLL" (ByVal f_v As Long) As Long ' 1=Frente 2=Verso 3=Ambas
Declare Function MC93_SetImageDirectory Lib "VIPS_DLL.DLL" (ByVal Diretorio As String) As Long ' Path para as Imagens
Declare Function MC93_Digitaliza Lib "VIPS_DLL.DLL" (ByVal NumInicial As Long, ByVal FileRetorno As String, ByVal Estacao As Long) As Long
Declare Function MC93_Reset Lib "VIPS_DLL.DLL" () As Long
Declare Function MC93_SetGaugePos Lib "VIPS_DLL.DLL" (ByVal nTop As Long, ByVal nLeft As Long) As Long
Declare Function MC93_SetImageType Lib "VIPS_DLL.DLL" (ByVal ImgType As Long) As Long ' 1=BMP 2=TIF 3=JPG
Declare Function MC93_SetQuality Lib "VIPS_DLL.DLL" (ByVal FatorQualidade As Long) As Long ' 30=Default
Declare Function MC93_CutBords Lib "VIPS_DLL.DLL" (ByVal T_F As Long) As Long  ' 1=Corta, 0=Não Corta
Declare Function MC93_SetCutBordsValue Lib "VIPS_DLL.DLL" (ByVal Valor As Long) As Long  ' Valor Baixo=Não corta  Valor Alto=Corta muito
                                                                                         
' Funcoes da dll VIPSDLL.Dll do Unibanco

' Inicializa as Dlls
Declare Function VIPS_Init Lib "VIPSDLL.DLL" () As Long
' Libera as Dlls
Declare Sub VIPS_Done Lib "VIPSDLL.DLL" ()
' Seleciona a porta serial (Default = 1)
Declare Sub VIPS_SetComPort Lib "VIPSDLL.DLL" (ByVal ComPort As Long)
' Seleciona a resolucao ( 100 ou 200 DPI ) (Default = 100)
Declare Sub VIPS_SetResolution Lib "VIPSDLL.DLL" (ByVal Res As Long)
' Seleciona leitora ( 1 = CMC7; 2 = Cod. Barras; 3 = Ambas ) (Default = 3)
Declare Sub VIPS_SetReader Lib "VIPSDLL.DLL" (ByVal Reader As Long)
' Seleciona a quantidade de Escaninhos ( Nro. de escaninhos deve ser impar )
' (Default = 1)
Declare Sub VIPS_SetBoxes Lib "VIPSDLL.DLL" (ByVal Boxes As Long)
' Seleciona a qtde maxima de docs por escanhinho (Default = 200)
Declare Sub VIPS_SetMaxDocBox Lib "VIPSDLL.DLL" (ByVal MaxDocBox As Long)
' Seleciona o escanhinho default ( Default = 0 : Nenhum )
Declare Sub VIPS_SetBoxDefault Lib "VIPSDLL.DLL" (ByVal BoxDefault As Long)
' Seleciona o tipo de imagem gerada ( 1 = BMP; 3 = JPG ) (Default = 3)
Declare Sub VIPS_SetImageType Lib "VIPSDLL.DLL" (ByVal ImageType As Long)
' Seleciona o fator de compressao do JPG ( Defualt = 30 )
Declare Sub VIPS_SetCompress Lib "VIPSDLL.DLL" (ByVal Fator As Long)
' Seleciona se deve cortar as bordas
' ( Valor-> 0 a 255, Zero nao corta. Quanto mais alto, mais corta )
' (Default = 75)
Declare Sub VIPS_SetCutBords Lib "VIPSDLL.DLL" (ByVal Valor As Long)
' Seleciona o diretorio onde serao gravadas as imagens
Declare Sub VIPS_SetImageDirectory Lib "VIPSDLL.DLL" (ByVal Diretorio As String)
' Seleciona o arq de conf. da camera ( nao precisa do caminho )
Declare Sub VIPS_SetCameraFile Lib "VIPSDLL.DLL" (ByVal CameraFile As String)
' Executa a captura
Declare Function VIPS_Captura Lib "VIPSDLL.DLL" (ByVal AgProc As Long, ByVal Lote As Long, ByVal ArqRetorno As String, ByVal Append As Long) As Long
' Executa a Recaptura
Declare Function VIPS_Recaptura Lib "VIPSDLL.DLL" (ByVal AgProc As Long, ByVal Lote As Long, ByVal SeqInic As Long, ByVal ArqRetorno As String, ByVal Append As Long) As Long
' Reseta as Dlls
Declare Function VIPS_Reset Lib "VIPSDLL.DLL" () As Long




