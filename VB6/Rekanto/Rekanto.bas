Attribute VB_Name = "Rekanto"
'Public db As New ADODB.Connection
Public DB As Connection


'Declaração das Funções da DLL da mini impressora
Public Declare Function AcionaGuilhotina Lib "mp2032.dll" (ByVal Modo As Integer) As Integer
Public Declare Function AutenticaDoc Lib "mp2032.dll" (ByVal BufTras As String, ByVal Tempo As Integer) As Integer
Public Declare Function BematechTX Lib "mp2032.dll" (ByVal BufTrans As String) As Integer
Public Declare Function CaracterGrafico Lib "mp2032.dll" (ByVal Buffer As String, ByVal TamBuffer As Integer) As Integer
Public Declare Function ComandoTX Lib "mp2032.dll" (ByVal BufTrans As String, ByVal TamBufTrans As Integer) As Integer
Public Declare Function ConfiguraModeloImpressora Lib "mp2032.dll" (ByVal ModeloImpressora As Integer) As Integer
Public Declare Function ConfiguraTamanhoExtrato Lib "mp2032.dll" (ByVal NumeroLinhas As Integer) As Integer
Public Declare Function DocumentInserted Lib "mp2032.dll" () As Integer
Public Declare Function EsperaImpressao Lib "mp2032.dll" () As Integer
Public Declare Function FechaPorta Lib "mp2032.dll" () As Integer
Public Declare Function FormataTX Lib "mp2032.dll" (ByVal BufTras As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal expand As Integer, ByVal enfat As Integer) As Integer
Public Declare Function HabilitaEsperaImpressao Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function HabilitaExtratoLongo Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function HabilitaPresenterRetratil Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function IniciaPorta Lib "mp2032.dll" (ByVal iPorta As String) As Integer
Public Declare Function Le_Status Lib "mp2032.dll" () As Integer
Public Declare Function Le_Status_Gaveta Lib "mp2032.dll" () As Integer
Public Declare Function ProgramaPresenterRetratil Lib "mp2032.dll" (ByVal Tempo As Integer) As Integer
Public Declare Function Status_Porta Lib "mp2032.dll" () As Integer
Public Declare Function VerificaPapelPresenter Lib "mp2032.dll" () As Integer

'Função para configuração dos códigos de barras
Public Declare Function ConfiguraCodigoBarras Lib "mp2032.dll" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer

'Funções para impressão do bitmap
Public Declare Function ImprimeBmpEspecial Lib "mp2032.dll" (ByVal FileName As String, ByVal xScale As Integer, ByVal yScale As Integer, ByVal angle As Integer) As Integer
Public Declare Function ImprimeBitmap Lib "mp2032.dll" (ByVal FileName As String, ByVal mode As Integer) As Integer
Public Declare Function AjustaLarguraPapel Lib "mp2032.dll" (ByVal width As Integer) As Integer
Public Declare Function SelectDithering Lib "mp2032.dll" (ByVal algorithm As Integer) As Integer

'Funções para impressão dos códigos de barras
Public Declare Function ImprimeCodigoBarrasUPCA Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasUPCE Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasEAN13 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasEAN8 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE39 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE93 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE128 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasITF Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODABAR Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasISBN Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasMSI Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasPLESSEY Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasPDF417 Lib "mp2032.dll" (ByVal NivelCorrecaoErros As Integer, ByVal Altura As Integer, ByVal Largura As Integer, ByVal Colunas As Integer, ByVal Codigo As String) As Integer
Public Sub Main()

    Dim x As Integer
    Dim Linha As String
    Dim Database As String

    f = FreeFile
    Open App.Path & "\rekanto.ini" For Input As #f
    Line Input #f, Linha
    Database = Mid(Linha, 10)
    Close #f

    'Abre a conexão de banco de dados
    Set DB = New Connection
    DB.CursorLocation = adUseClient
    DB.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database

    'DB.Open "Rekanto", "", ""

    Principal.Show 1
End Sub
