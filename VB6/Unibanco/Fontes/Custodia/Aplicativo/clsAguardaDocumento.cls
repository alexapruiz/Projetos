VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AguardaDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''
'Exemplo da chamada à esta classe'
''''''''''''''''''''''''''''''''''

'    Dim x As New AguardaDocumento
'    Dim Proc_Selecionar         As New Custodia.Selecionar
'
'    x.SetConnection g_cMainConnection
'    x.PesquisaStatus = "5"
'    x.Tempo = 2
'    x.SetStatusBar = Me.StatusBar1
'    x.SetProgressBar Me.ProgressBar1
'    x.SQL = Proc_Selecionar.GetSupervisor(m_DataProcessamento, g_Parametros.TMP_Pendente)
'    x.SetStaus "Aguardando novo borderô"
'
'
'    Do While x.Espera()
'        DoEvents
'        x.SQL = Proc_Selecionar.GetSupervisor(m_DataProcessamento, g_Parametros.TMP_Pendente)
'    Loop
'
'    Set x = Nothing

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


Private m_Tempo             As Long
Private m_SQL               As String
Private m_Connection        As New ADODB.Connection
Private m_Recordset         As New ADODB.Recordset
Private m_ProgressBar       As Object 'ProgressBar
Private m_StatusBar         As Object 'StatusBar
Private m_Finalizar         As Boolean
Private m_Finalizado        As Boolean
Public Function ExisteDocumento() As Boolean

    Dim lStart      As Long
    Dim rst         As New ADODB.Recordset
    Dim i           As Integer
    Dim iProgress   As Integer
    
    ExisteDocumento = True
    
    If m_ProgressBar Is Nothing Then
        MsgBox "ProgressBar é obrigatório.", vbExclamation
        Exit Function
    End If
    
    If m_StatusBar Is Nothing Then
        MsgBox "StatusBar é obrigatório.", vbExclamation
        Exit Function
    End If

    lStart = Int(Timer)
    
    iProgress = 1
    'm_ProgressBar.Value = iProgress

    Do While (Abs(lStart - Timer) < m_Tempo And m_Finalizar = False)
        ''''''''''''''''''''''''''
        'Incrementa o ProgressBar'
        ''''''''''''''''''''''''''
        'iProgress = Int((Abs(lStart - Timer) / m_Tempo) * 100)
        iProgress = (Abs(lStart - Int(Timer)) / m_Tempo) * 100
        
        If iProgress <= 100 Then
            m_ProgressBar.Value = iProgress
        End If
        
        DoEvents
    Loop
    
    If m_Finalizar Then
        Set m_Recordset = Nothing
'        Set m_StatusBar = Nothing
'        Set m_ProgressBar = Nothing
        m_Finalizado = True
        Exit Function
    End If
    
    m_ProgressBar.Value = 0
    Set rst = m_Connection.Execute(m_SQL)
    
    If rst.RecordCount = 0 Then
        ExisteDocumento = False
        Set m_Recordset = Nothing
'        Set m_StatusBar = Nothing
'        Set m_ProgressBar = Nothing
    Else
        Set m_Recordset = rst
'        Set m_ProgressBar = Nothing
        Me.SetStatus ""
'        Set m_StatusBar = Nothing
    End If
    
    
    
End Function


Public Property Get Finalizado() As Boolean
    Finalizado = m_Finalizado
End Property

Public Sub Finalizar()
    m_Finalizar = True
End Sub

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = m_Recordset
End Property

Public Sub SetConnection(ByVal pConnection As ADODB.Connection)
    Set m_Connection = pConnection
End Sub

Public Sub SetProgressBar(ByRef pProgressBar As ProgressBar)
    Set m_ProgressBar = pProgressBar
End Sub

Public Sub SetStatus(ByVal pStatus As String)

    If m_StatusBar Is Nothing Then
        MsgBox "StatusBar é obrigatório.", vbExclamation
        Exit Sub
    End If
    
    m_StatusBar.Panels(1).Text = pStatus
    
End Sub

Public Sub SetStatusBar(ByRef pStatusBar As StatusBar)
    Set m_StatusBar = pStatusBar
End Sub

Public Property Let SQL(ByVal pSQL As String)
    m_SQL = pSQL
End Property

Public Property Let Tempo(ByVal pTempo As Long)
    m_Tempo = pTempo
End Property

Public Property Get Tempo() As Long
    Tempo = m_Tempo
End Property


Private Sub Class_Initialize()

    m_Finalizar = False
    m_Finalizado = False
    
    ''''''''''''''''''''''''
    'Default de 20 segundos'
    ''''''''''''''''''''''''
    m_Tempo = 20
End Sub





Private Sub Class_Terminate()

    Set m_StatusBar = Nothing
    Set m_ProgressBar = Nothing

End Sub


