VERSION 5.00
Begin VB.UserControl Scanner 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   576
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   576
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   InvisibleAtRuntime=   -1  'True
   KeyPreview      =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   576
   ScaleWidth      =   576
   ToolboxBitmap   =   "Scanner.ctx":0000
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   384
      Left            =   96
      Picture         =   "Scanner.ctx":0312
      Top             =   96
      Width           =   384
   End
End
Attribute VB_Name = "Scanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Dim HwndLibLA As Long
 Dim HwndLib9X As Long
 Dim HwndLibNT As Long
  
'Propriedades do Sistema Operacional
 Public Enum enumWINVERSION
    eWIN_ER = 0     'Não foi possivel obter a versao do sistema operacional
    eWIN_NT = 1     'Windows NT
    eWIN_9X = 2     'Windows 95, 98
 End Enum
 
'Valores de Retorno da leitura
 Public Enum enumRetorno
    eFIM = 0     'Final de cheques no Alimentador
    eOK = 1      'Leitura Terminada OK
    eEsc = 2     'Teclou Esc Durante leitura
    eFalha = 3   'Falha na Leitura
    eTimeOut = 4 'Time Out
    eErro = 9    'Erro no Módulo
 End Enum
 
'Enun Scanner's
 Public Enum enumScanner
    eNulo = 0
    eL100 = 1
    eLA93 = 2
 End Enum
 
'Default Property Values:
 Const m_def_BaudRate = 2400
 Const m_def_CommPort = 1
 Const m_def_wordLenght = 7
 Const m_def_Parity = 0
 Const m_def_StopBits = 1
 Const m_def_OS = 2
 Const m_def_Scanner = 0
 Const m_def_Habilitado = False
 
'Property Variables:
 Dim m_BaudRate As Long
 Dim m_CommPort As Integer
 Dim m_wordLenght As Integer
 Dim m_Parity As Integer
 Dim m_StopBits As Integer
 Dim m_OS As Integer
 Dim m_Scanner As Integer
 Dim m_Habilitado As Boolean
   
'Propriedades de Leitura
 Dim m_CMC7_Campo1 As String * 8
 Dim m_CMC7_Campo2 As String * 10
 Dim m_CMC7_Campo3 As String * 12
 
 Dim M_RunError As ErrObject
 
'Event Declarations:
 Event WriteProperties(PropBag As PropertyBag)
Private Sub UserControl_InitProperties()

    m_BaudRate = m_def_BaudRate
    m_CommPort = m_def_CommPort
    m_wordLenght = m_def_wordLenght
    m_Parity = m_def_Parity
    m_StopBits = m_def_StopBits
    m_OS = m_def_OS
    m_Scanner = m_def_Scanner
    m_Habilitado = m_def_Habilitado
     
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BaudRate = PropBag.ReadProperty("BaudRate", m_def_BaudRate)
    m_CommPort = PropBag.ReadProperty("CommPort ", m_def_CommPort)
    m_wordLenght = PropBag.ReadProperty("wordLenght", m_def_wordLenght)
    m_Parity = PropBag.ReadProperty("Parity", m_def_Parity)
    m_StopBits = PropBag.ReadProperty("StopBits", m_def_StopBits)
    m_OS = PropBag.ReadProperty("OS", m_def_OS)
    m_Scanner = PropBag.ReadProperty("Scanner", m_def_Scanner)
    m_Habilitado = PropBag.ReadProperty("Habilitado", m_def_Habilitado)
    
End Sub
Private Sub UserControl_Resize()

    UserControl.Width = 550
    UserControl.Height = 500

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BaudRate", m_BaudRate, m_def_BaudRate)
    Call PropBag.WriteProperty("CommPort", m_CommPort, m_def_CommPort)
    Call PropBag.WriteProperty("wordLenght", m_wordLenght, m_def_BaudRate)
    Call PropBag.WriteProperty("Parity", m_Parity, m_def_Parity)
    Call PropBag.WriteProperty("StopBits", m_StopBits, m_def_StopBits)
    Call PropBag.WriteProperty("OS", m_OS, m_def_OS)
    Call PropBag.WriteProperty("Scanner", m_Scanner, m_def_Scanner)
    Call PropBag.WriteProperty("Habilitado", m_Habilitado, m_def_Habilitado)
    
End Sub
Public Property Get Habilitado() As Boolean
    Habilitado = m_Habilitado
End Property
Public Property Let Habilitado(ByVal New_Habilitado As Boolean)
    m_Habilitado = New_Habilitado
    PropertyChanged "Habilitado"
End Property
Public Property Get Scanner() As enumScanner
    Scanner = m_Scanner
End Property
Public Property Let Scanner(ByVal New_Scanner As enumScanner)
    m_Scanner = New_Scanner
    PropertyChanged "Scanner"
End Property
Public Property Get MSGLEFT() As Integer
    MSGLEFT = m_MSGLEFT
End Property
Public Property Let MSGLEFT(ByVal New_MSGLEFT As Integer)
    m_MSGLEFT = New_MSGLEFT
End Property
Public Property Get MSGTOP() As Integer
    MSGTOP = m_MSGTOP
End Property
Public Property Let MSGTOP(ByVal New_MSGTOP As Integer)
    m_MSGTOP = New_MSGTOP
End Property
Public Property Let CommPort(ByVal New_CommPort As Integer)
    m_CommPort = New_CommPort
    PropertyChanged "CommPort"
End Property
Public Property Get CommPort() As Integer
    CommPort = m_CommPort
End Property
Public Property Let WordLenght(ByVal New_WordLenght As Integer)
    m_wordLenght = New_WordLenght
    PropertyChanged "WordLenght"
End Property
Public Property Get WordLenght() As Integer
    WordLenght = m_wordLenght
End Property
Public Property Let Parity(ByVal New_Parity As Integer)
    m_Parity = New_Parity
    PropertyChanged "Parity"
End Property
Public Property Get Parity() As Integer
    Parity = m_Parity
End Property
Public Property Let StopBits(ByVal New_StopBits As Integer)
    m_StopBits = New_StopBits
    PropertyChanged "StopBits"
End Property
Public Property Get StopBits() As Integer
    StopBits = m_StopBits
End Property
Public Property Get OS() As enumWINVERSION
    OS = m_OS
End Property
Public Property Let OS(ByVal New_OS As enumWINVERSION)
    m_OS = New_OS
    PropertyChanged "OS"
End Property
Public Property Get BaudRate() As Long
    BaudRate = m_BaudRate
End Property
Public Property Let BaudRate(ByVal New_BaudRate As Long)
    m_BaudRate = New_BaudRate
    PropertyChanged "BaudRate"
End Property
Public Property Get CMC7_Campo1() As String
    CMC7_Campo1 = m_CMC7_Campo1
End Property
Public Property Get CMC7_Campo2() As String
    CMC7_Campo2 = m_CMC7_Campo2
End Property
Public Property Get CMC7_Campo3() As String
    CMC7_Campo3 = m_CMC7_Campo3
End Property
Public Property Get Erro() As ErrObject
    Set Erro = M_RunError
End Property
Public Function Inicializa() As Boolean
    On Error GoTo Erro
    
    Dim Ret, i As Integer
        
    If Me.OS = eWIN_NT And Me.Scanner = eL100 Then
        HwndLibNT = LoadLibrary("DTC32NT.DLL")
        For i = 1 To 6
            'Parametro de configuracao
             Select Case i
                 Case 1
                     Ret = DTCNT_SetCommPort(Me.CommPort)
                 Case 2
                     Ret = DTCNT_SetBaudRate(Me.BaudRate)
                 Case 3
                     Ret = DTCNT_SetWordLength(Me.WordLenght)
                 Case 4
                     Ret = DTCNT_SetParity(Me.Parity)
                 Case 5
                     Ret = DTCNT_SetStopBits(Me.StopBits)
                 Case 6
                     Ret = DTCNT_Init()
             End Select
        
            If Ret <> 1 Then
                  Set M_RunError = Err
                  M_RunError.Number = 901
                  M_RunError.Description = "Retorno indevido na Chamada de Aplicativo da L100 (DLL-Módulo Inicilização) - Função: " & Str(i)
                  GoTo Erro
            End If
        Next
          
          Inicializa = True
          
    ElseIf Me.OS = eWIN_9X And Me.Scanner = eL100 Then
        HwndLib9X = LoadLibrary("DTC329X.DLL")
        For i = 1 To 6
              'Parametro de configuracao
               Select Case i
                   Case 1
                       Ret = DTC9X_SetCommPort(Me.CommPort)
                   Case 2
                       Ret = DTC9X_SetBaudRate(Me.BaudRate)
                   Case 3
                       Ret = DTC9X_SetWordLength(Me.WordLenght)
                   Case 4
                       Ret = DTC9X_SetParity(Me.Parity)
                   Case 5
                       Ret = DTC9X_SetStopBits(Me.StopBits)
                   Case 6
                       Ret = DTC9X_Init()
               End Select
          
              If Ret <> 1 Then
                    Set M_RunError = Err
                    M_RunError.Number = 901
                    M_RunError.Description = "Retorno indevido na Chamada de Aplicativo da L100 (DLL-Módulo Inicilização) - Função: " & Str(i)
                    GoTo Erro
              End If
          Next
          
          Inicializa = True
          
    ElseIf Me.Scanner = eLA93 Then
        HwndLibLA = LoadLibrary("La93.dll")
        For i = 1 To 3
              'Parametro de configuracao
               Select Case i
                   Case 1
                       Ret = LA93_SetCommPort(Me.CommPort)
                   Case 2
                       Ret = LA93_SetNumEscaninhos(1)
                   Case 3
                       Ret = LA93_Init()
               End Select
          
              If Ret <> 1 Then
                    Set M_RunError = Err
                    M_RunError.Number = 901
                    M_RunError.Description = "Retorno indevido na Chamada de Aplicativo da LA93" & vbCrLf & "(DLL-Módulo Inicialização) - Função: " & Str(i)
                    GoTo Erro
              End If
        Next
        
        Inicializa = True
    
    Else
        M_RunError.Number = 901
        M_RunError.Description = "Propriedade(s) do Scanner Control inválido(s) - (Módulo Inicialização) "
        GoTo Erro
    End If
    
Exit Function

Erro:
    m_Inicializa = False
    
End Function
Public Function Le() As enumRetorno
On Error GoTo Erro:

Dim CMC7 As String * 128
Dim strCMC7 As String * 128
Dim clearCMC7 As String * 30
Dim Ret As Integer
    

clearCMC7 = ""
m_CMC7_Campo1 = ""
m_CMC7_Campo2 = ""
m_CMC7_Campo3 = ""

If Me.Scanner = eL100 Then
    If Me.OS = eWIN_NT Then
        DoEvents
       'Starta Leitura
        Ret = DTCNT_Read(CMC7)
    ElseIf Me.OS = eWIN_9X Then
        DoEvents
       'Starta Leitura
        Ret = DTC9X_Read(CMC7)
    End If
    
    DoEvents
    
   'Pega apenas caracteres validos.
    CMC7 = Replace(CMC7, " ", "", 1, Len(CMC7), vbTextCompare)
    clearCMC7 = Mid(Trim(CMC7), 2, 8) & Mid(Trim(CMC7), 11, 10) & Mid(Trim(CMC7), 22, 12)
    
    If (Ret > 0 And Len(Trim(Ret)) > 1) And _
       (InStr(1, clearCMC7, "!", vbTextCompare) = 0 And _
        InStr(1, clearCMC7, "=", vbTextCompare) = 0 And _
        InStr(1, clearCMC7, ">", vbTextCompare) = 0 And _
        InStr(1, clearCMC7, "?", vbTextCompare) = 0 And _
        InStr(1, clearCMC7, ":", vbTextCompare) = 0) Then
        
        m_CMC7_Campo1 = Mid(Trim(CMC7), 2, 8)
        m_CMC7_Campo2 = Mid(Trim(CMC7), 11, 10)
        m_CMC7_Campo3 = Mid(Trim(CMC7), 22, 12)
        
       'Se leitura bem sucedida
        Le = eOK
        
    ElseIf Ret <> 0 And Ret <> -2 Then
           'Falha na Leitura
            Le = eFalha
    ElseIf Ret = 0 Then
       'O usuario teclou (Esc) durante a leitura
        Le = eEsc
    ElseIf Ret = -2 Then
       'Saiu por timeout
        Le = eTimeOut
    End If
ElseIf Me.Scanner = eLA93 Then
        DoEvents
        
       'Starta Leitura
       ' Ret = LA93_Read(m_CMC7_Campo1, m_CMC7_Campo2, m_CMC7_Campo3)
        
        Ret = LA_ClearBuffer()
        Ret = LA93_Read(strCMC7)
                
        
        ' Tratamento da Leitura
        
         strCMC7 = Replace(strCMC7, " ", "", 1, Len(strCMC7), vbTextCompare)
         strCMC7 = Replace(strCMC7, "!", "", 1, Len(strCMC7), vbTextCompare)
         strCMC7 = Replace(strCMC7, "=", "", 1, Len(strCMC7), vbTextCompare)
         strCMC7 = Replace(strCMC7, ">", "", 1, Len(strCMC7), vbTextCompare)
         strCMC7 = Replace(strCMC7, "<", "", 1, Len(strCMC7), vbTextCompare)
         strCMC7 = Replace(strCMC7, "?", "", 1, Len(strCMC7), vbTextCompare)
         strCMC7 = Replace(strCMC7, ":", "", 1, Len(strCMC7), vbTextCompare)
         
         strCMC7 = Mid(Trim(strCMC7), 1, 8) & Mid(Trim(strCMC7), 9, 10) & Mid(Trim(strCMC7), 19, 12)
    
        ' Fim Tratamento da leitura
        
        
         m_CMC7_Campo1 = Mid(Trim(strCMC7), 1, 8)
         m_CMC7_Campo2 = Mid(Trim(strCMC7), 9, 10)
         m_CMC7_Campo3 = Mid(Trim(strCMC7), 19, 12)
        
         clearCMC7 = m_CMC7_Campo1 & m_CMC7_Campo2 & m_CMC7_Campo3
        
        If Ret = 1 And IsNumeric(clearCMC7) Then
           'Se leitura bem sucedida
            Le = eOK
        ElseIf Ret = -2 Then
            Le = eFIM
        Else
           'Se erro na Leitura
            clearCMC7 = ""
            m_CMC7_Campo1 = ""
            m_CMC7_Campo2 = ""
            m_CMC7_Campo3 = ""

            Le = eFalha
        End If
      
Else
    GoTo Erro
End If

Exit Function

Erro:
    Le = eErro
    Set M_RunError = Err
    M_RunError.Number = 904
    M_RunError.Description = "Erro no aplicativo de comunicação com Scanner (módulo de Leitura)"
    
End Function
Public Function Finaliza() As Boolean

On Error GoTo Erro:

    If Me.OS = eWIN_NT And Me.Scanner = eL100 Then
        DTCNT_DeInit
        FreeLibrary (HwndLibNT)
    ElseIf Me.OS = eWIN_9X And Me.Scanner = eL100 Then
        DTC9X_DeInit
        FreeLibrary (HwndLib9X)
    ElseIf Me.Scanner = eLA93 Then
        LA93_DeInit
        FreeLibrary (HwndLibLA)
    Else
        GoTo Erro
    End If
    
    Finaliza = True
    
Exit Function

Erro:
    Set M_RunError = Err
    M_RunError.Number = 904
    M_RunError.Description = "Erro no aplicativo de comunicação com Scanner (módulo Finalização)"

End Function
Public Sub Eject()
On Error GoTo Erro:

    LA93_Eject (1)
    Exit Sub

Erro:
    Set M_RunError = Err
    M_RunError.Number = 904
    M_RunError.Description = "Erro no aplicativo de comunicação com Scanner (módulo Eject)"
    
End Sub

