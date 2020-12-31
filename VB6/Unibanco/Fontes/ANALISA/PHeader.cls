VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCabecalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Titulo                As String
Dim m_DataProcessamento     As Long
Dim m_Estacao               As String

Public Property Let DataProcessamento(ByVal pDataProcessamento As Long)
    m_DataProcessamento = pDataProcessamento
End Property


Public Property Let Estacao(ByVal pEstacao As String)
    m_Estacao = pEstacao
End Property

Public Sub Imprimir()

    Dim pLogo                   As IPictureDisp
    Dim FontSize                As Integer
    Dim sDataProcessamento      As String
    Dim sEstacao                As String
    Dim y                       As Long
    Dim sFont                   As String
    
    sFont = Printer.Font.Name
    Printer.Font.Name = "Times New Roman"
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    
    Set pLogo = LoadPicture(App.Path & "\Logo-Unibanco.jpg")
    
    FontSize = Printer.FontSize
    
    ''''''''''''''''
    'Imprime o Logo'
    ''''''''''''''''
    Printer.PaintPicture pLogo, 0, 0
    
    '''''''''''''''''''''''''''''''''
    'Imprime a data de Processamento'
    '''''''''''''''''''''''''''''''''
    y = Printer.CurrentY
    Printer.CurrentX = Printer.ScaleWidth - 2200
    Printer.Print "Data de Captura :"
    
    Printer.CurrentY = y
    
    sDataProcessamento = Format(Format(m_DataProcessamento, "0000-00-00"), "dd/mm/yyyy")
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(sDataProcessamento) - 50
    Printer.Print sDataProcessamento
    '''''''''''''''''''''''''''
    'Imprime o nome da estação'
    '''''''''''''''''''''''''''
    y = Printer.CurrentY
    Printer.CurrentX = Printer.ScaleWidth - 2200
    Printer.Print "Estação :"
    
    Printer.CurrentY = y
    
    sEstacao = m_Estacao
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(sEstacao) - 50
    Printer.Print sEstacao
    
    ''''''''''''''''''''''
    'Imprime a data atual'
    ''''''''''''''''''''''
    y = Printer.CurrentY
    Printer.CurrentX = Printer.ScaleWidth - 2200
    Printer.Print "Data :"
    
    Printer.CurrentY = y
    
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Date) - 50
    Printer.Print Date
    
    ''''''''''''''''''
    'Imprime o Titulo'
    ''''''''''''''''''
    Printer.CurrentY = 200
    Printer.FontSize = 14
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(m_Titulo)) / 2
    Printer.Print m_Titulo
    
    Printer.FontSize = FontSize
    
    '''''''''''''''''''''''''''''''''''''''
    'Imprime uma linha no fim do cabecalho'
    '''''''''''''''''''''''''''''''''''''''
    Printer.CurrentY = Printer.CurrentY + 200
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY + 15), , BF
    
    Printer.Font.Name = sFont

End Sub

Public Property Let Titulo(ByVal pTitulo As String)
    m_Titulo = pTitulo
End Property


