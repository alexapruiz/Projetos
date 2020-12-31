VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   648
   ClientLeft      =   2304
   ClientTop       =   3540
   ClientWidth     =   1356
   LinkTopic       =   "Form1"
   ScaleHeight     =   648
   ScaleWidth      =   1356
   Visible         =   0   'False
   Begin Crystal.CrystalReport CrRelEstatisca 
      Left            =   60
      Top             =   72
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   262150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private qryGetResumoQuantidades As rdoQuery            ' query Contador de Documentos
Private RsAux                   As rdoResultset        ' Recordset

'* 1 - Capas s� com Pagamentos
'* 2 - Capas s� com Dep�sitos
'* 3 - Capas Com Pagto/Deposito (Misto)
'* 4 - Vari�veis Auxiliares - Contadores de Loop

' * 1
Dim CppagQtdeCapMal        As Integer 'Capas s� Com Pagamento Qtde Malotes
Dim CppagQtdeCapEnv        As Integer 'Capas s� Com Pagamento Qtde Envelope
Dim CppagQtdePagMal        As Integer 'Capas s� Com Pagamento Qtde Pagto Malote
Dim CppagQtdePagEnv        As Integer 'Capas s� Com Pagamento Qtde Pagto Envelope
Dim CppagQtdeChPagMal      As Integer 'Capas s� Com Pagamento Qtde Ch Pagto Malote
Dim CppagQtdeChPagEnv      As Integer 'Capas s� Com Pagamento Qtde Ch Pagto Envelope
Dim CppagQtdeDepMal        As String  'Capas s� Com Pagamento Qtde Deposito Malote
Dim CppagQtdeDepEnv        As String  'Capas s� Com Pagamento Qtde Deposito Envelope
Dim CppagQtdeChDepMal      As String  'Capas s� Com Pagamento Qtde ch Deposito Malote
Dim CppagQtdeChDepEnv      As String  'Capas s� Com Pagamento Qtde Ch Deposito Envelope
' * 2
Dim CpDepQtdeCapMal        As Integer 'Capas s� com Dep�sitos Qtde Malote
Dim CpDepQtdeCapEnv        As Integer 'Capas s� com Dep�sitos Qtde Envelope
Dim CpDepQtdePagMal        As String  'Capas s� com Dep�sitos Qtde Pagto Malote
Dim CpDepQtdePagEnv        As String  'Capas s� com Dep�sitos Qtde Pagto Envelope
Dim CpDepQtdeChPagMal      As String  'Capas s� com Dep�sitos Qtde Ch Pagto Malote
Dim CpDepQtdeChPagEnv      As String  'Capas s� com Dep�sitos Qtde Ch Pagto Envelope
Dim CpDepQtdeDepMal        As Integer 'Capas s� com Dep�sitos Qtde Dep�sito Malote
Dim CpDepQtdeDepEnv        As Integer 'Capas s� com Dep�sitos Qtde Dep�sito Envelope
Dim CpDepQtdeChDepMal      As Integer 'Capas s� com Dep�sitos Qtde Ch Dep�sito Malote
Dim CpDepQtdeChDepEnv      As Integer 'Capas s� com Dep�sitos Qtde Ch Dep�sito Envelope
' * 3
Dim CpPagDepQtdeCapMal     As Integer 'Capas com Pagto/Dep�sito Qtde Malote
Dim CpPagDepQtdeCapEnv     As Integer 'Capas com Pagto/Dep�sito Qtde Envelope
Dim CpPagDepQtdePagMal     As Integer 'Capas com Pagto/Dep�sito Qtde Pagto Malote
Dim CpPagDepQtdePagEnv     As Integer 'Capas com Pagto/Dep�sito Qtde Pagto Envelope
Dim CpPagDepQtdeChPagMal   As Integer 'Capas com Pagto/Dep�sito Qtde Ch Pagto Malote
Dim CpPagDepQtdeChPagEnv   As Integer 'Capas com Pagto/Dep�sito Qtde Ch Pagto Envelope
Dim CpPagDepQtdeDepMal     As Integer 'Capas com Pagto/Dep�sito Qtde Dep�sito Malote
Dim CpPagDepQtdeDepEnv     As Integer 'Capas com Pagto/Dep�sito Qtde Dep�sito Envelope
Dim CpPagDepQtdeChDepMal   As Integer 'Capas com Pagto/Dep�sito Qtde Ch Dep�sito Malote
Dim CpPagDepQtdeChDepEnv   As Integer 'Capas com Pagto/Dep�sito Qtde Ch Dep�sito Envelope
' * 4
Dim CCpPagtoQtdeCapa        As Integer 'Contador de Loop Capas s� com Pagto Qtde Capa
Dim CCpPagtoQtdePagto       As Integer 'Contador de Loop Capas s� com Pagto Qtde Pgto
Dim CcpPagtoQtdeChPagto     As Integer 'Contador de Loop Capas s� com Pagto Qtde Ch Pagto
Dim CCpDepQtdeCapa          As Integer 'Contador de Loop Capas s� Com Dep�sito Qtde Capa
Dim CcpDepQtdeDep           As Integer 'Contador de Loop Capas s� Com Dep�sito Qtde Dep�sito
Dim CcpDepQtdeChDep         As Integer 'Contador de Loop Capas s� Com Dep�sito Qtde Ch Dep�sito
Dim CCpPagtoDepQtdeCapa     As Integer 'Contador de Loop Capas com Pagto/Dep�sito Qtde Capa
Dim CCpPagtoDepQtdePagto    As Integer 'Contador de Loop Capas com Pagto/Dep�sito Qtde Pagto
Dim CcpPagtoDepQtdeChPagto  As Integer 'Contador de Loop Capas com Pagto/Dep�sito Qtde Ch Pagto
Dim CcpPagtoDepQtdeDep      As Integer 'Contador de Loop Capas com Pagto/Dep�sito Qtde Dep�sito
Dim CcpPagtoDepQtdeChDep    As Integer 'Contador de Loop Capas com Pagto/Dep�sito Qtde Ch Dep�sito

Private Sub Form_Activate()

'Valor default para  as vari�veis(Zero)

'* 1 - Capas s� com Pagamentos
'* 2 - Capas s� com Dep�sitos
'* 3 - Capas Com Pagto/Deposito (Misto)

' * 1
CppagQtdeCapMal = 0         'Capas s� Com Pagamento Qtde Malotes
CppagQtdeCapEnv = 0         'Capas s� Com Pagamento Qtde Envelope
CppagQtdePagMal = 0         'Capas s� Com Pagamento Qtde Pagto Malote
CppagQtdePagEnv = 0         'Capas s� Com Pagamento Qtde Pagto Envelope
CppagQtdeChPagMal = 0       'Capas s� Com Pagamento Qtde Ch Pagto Malote
CppagQtdeChPagEnv = 0       'Capas s� Com Pagamento Qtde Ch Pagto Envelope
CppagQtdeDepMal = "-"       'Capas s� Com Pagamento Qtde Deposito Malote
CppagQtdeDepEnv = "-"       'Capas s� Com Pagamento Qtde Deposito Envelope
CppagQtdeChDepMal = "-"     'Capas s� Com Pagamento Qtde ch Deposito Malote
CppagQtdeChDepEnv = "-"     'Capas s� Com Pagamento Qtde Ch Deposito Envelope
' * 2
CpDepQtdeCapMal = 0         'Capas s� com Dep�sitos Qtde Malote
CpDepQtdeCapEnv = 0         'Capas s� com Dep�sitos Qtde Envelope
CpDepQtdePagMal = "-"       'Capas s� com Dep�sitos Qtde Pagto Malote
CpDepQtdePagEnv = "-"       'Capas s� com Dep�sitos Qtde Pagto Envelope
CpDepQtdeChPagMal = "-"     'Capas s� com Dep�sitos Qtde Ch Pagto Malote
CpDepQtdeChPagEnv = "-"     'Capas s� com Dep�sitos Qtde Ch Pagto Envelope
CpDepQtdeDepMal = 0         'Capas s� com Dep�sitos Qtde Dep�sito Malote
CpDepQtdeDepEnv = 0         'Capas s� com Dep�sitos Qtde Dep�sito Envelope
CpDepQtdeChDepMal = 0       'Capas s� com Dep�sitos Qtde Ch Dep�sito Malote
CpDepQtdeChDepEnv = 0       'Capas s� com Dep�sitos Qtde Ch Dep�sito Envelope
' * 3
CpPagDepQtdeCapMal = 0      'Capas com Pagto/Dep�sito Qtde Malote
CpPagDepQtdeCapEnv = 0      'Capas com Pagto/Dep�sito Qtde Envelope
CpPagDepQtdePagMal = 0      'Capas com Pagto/Dep�sito Qtde Pagto Malote
CpPagDepQtdePagEnv = 0      'Capas com Pagto/Dep�sito Qtde Pagto Envelope
CpPagDepQtdeChPagMal = 0    'Capas com Pagto/Dep�sito Qtde Ch Pagto Malote
CpPagDepQtdeChPagEnv = 0    'Capas com Pagto/Dep�sito Qtde Ch Pagto Envelope
CpPagDepQtdeDepMal = 0      'Capas com Pagto/Dep�sito Qtde Dep�sito Malote
CpPagDepQtdeDepEnv = 0      'Capas com Pagto/Dep�sito Qtde Dep�sito Envelope
CpPagDepQtdeChDepMal = 0    'Capas com Pagto/Dep�sito Qtde Ch Dep�sito Malote
CpPagDepQtdeChDepEnv = 0    'Capas com Pagto/Dep�sito Qtde Ch Dep�sito Envelope

End Sub
Private Sub Form_Load()

Set qryGetResumoQuantidades = Geral.Banco.CreateQuery("", "{call GetResumoQuantidades (?)}")

    With qryGetResumoQuantidades
        .rdoParameters(0).Value = Geral.DataProcessamento
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    '/Valores Default
    CppagQtdeDepMal = "-"
    CppagQtdeDepEnv = "-"
    CppagQtdeChDepMal = "-"
    CppagQtdeChDepEnv = "-"
    CpDepQtdePagMal = "-"
    CpDepQtdePagEnv = "-"
    CpDepQtdeChPagMal = "-"
    CpDepQtdeChPagEnv = "-"
    '/Valores Default

    If Not RsAux.EOF Then

        Call CpPagtoQtdeCapa
        Call CpPagtoQtdePagto
        Call CpPagtoQtdeChPagto
        Call CpDepQtdeCapa
        Call CpDepQtdeDep
        Call CpDepQtdeChDep
        Call CpPagtoDepQtdeCapa
        Call CpPagtoDepQtdePagto
        Call CpPagtoDepQtdeChPagto
        Call CpPagtoDepQtdeDep
        Call CpPagtoDepQtdeChDep
        Call PreparaRelEstatistica
    End If
                
End Sub
Public Sub CpPagtoQtdeCapa()

    For CCpPagtoQtdeCapa = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdeCapEnv = RsAux!qtde
        Else
           CppagQtdeCapMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults

End Sub
Public Sub CpPagtoQtdePagto()

    For CCpPagtoQtdePagto = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdePagEnv = RsAux!qtde
        Else
           CppagQtdePagMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults
    
End Sub
Public Sub CpPagtoQtdeChPagto()

    For CcpPagtoQtdeChPagto = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdeChPagEnv = RsAux!qtde
        Else
           CppagQtdeChPagMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults

End Sub
Public Sub CpDepQtdeCapa()

    For CCpDepQtdeCapa = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeCapEnv = RsAux!qtde
        Else
           CpDepQtdeCapMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults


End Sub
Public Sub CpDepQtdeDep()


    For CcpDepQtdeDep = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeDepEnv = RsAux!qtde
        Else
           CpDepQtdeDepMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults

End Sub
Public Sub CpDepQtdeChDep()

    For CcpDepQtdeChDep = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeChDepEnv = RsAux!qtde
        Else
           CpDepQtdeChDepMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults

End Sub
Public Sub CpPagtoDepQtdeCapa()

    For CCpPagtoDepQtdeCapa = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeCapEnv = RsAux!qtde
        Else
           CpPagDepQtdeCapMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults

End Sub
Public Sub CpPagtoDepQtdePagto()

    For CCpPagtoDepQtdePagto = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdePagEnv = RsAux!qtde
        Else
           CpPagDepQtdePagMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults

End Sub
Public Sub CpPagtoDepQtdeChPagto()

    For CcpPagtoDepQtdeChPagto = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeChPagEnv = RsAux!qtde
        Else
           CpPagDepQtdeChPagMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults
    
End Sub

Public Sub CpPagtoDepQtdeDep()

    For CcpPagtoDepQtdeDep = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeDepEnv = RsAux!qtde
        Else
           CpPagDepQtdeDepMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults
    
End Sub
Public Sub CpPagtoDepQtdeChDep()

    For CcpPagtoDepQtdeChDep = 0 To RsAux.RowCount - 1
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeChDepEnv = RsAux!qtde
        Else
           CpPagDepQtdeChDepMal = RsAux!qtde
        End If
    
        RsAux.MoveNext
        
    Next

    RsAux.MoreResults
    
End Sub

Public Sub PreparaRelEstatistica()

Dim DataFormatada As String

    DataFormatada = Mid(Geral.DataProcessamento, 7, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Mid(Geral.DataProcessamento, 1, 4)

    CrRelEstatisca.ReportFileName = App.Path & "\estatistica.rpt "
    
    CrRelEstatisca.Formulas(0) = "CpPgtoCapa    = '" & CppagQtdeCapMal & "'"
    CrRelEstatisca.Formulas(1) = "CpPgtoPgto    = '" & CppagQtdePagMal & "'"
    CrRelEstatisca.Formulas(2) = "CpPgtoChPagto = '" & CppagQtdeChPagMal & "'"
    CrRelEstatisca.Formulas(3) = "CpPgtoDep     = '" & CppagQtdeDepMal & "'"
    CrRelEstatisca.Formulas(4) = "CpPgtoChDep   = '" & CppagQtdeChDepMal & "'"

    CrRelEstatisca.Formulas(5) = "CpDepCapa     = '" & CpDepQtdeCapMal & "'"
    CrRelEstatisca.Formulas(6) = "CpDepPgto     = '" & CpDepQtdePagMal & "'"
    CrRelEstatisca.Formulas(7) = "CpDepChPgto   = '" & CpDepQtdeChPagMal & "'"
    CrRelEstatisca.Formulas(8) = "CpDepDep      = '" & CpDepQtdeDepMal & "'"
    CrRelEstatisca.Formulas(9) = "CpDepChDep    = '" & CpDepQtdeChDepMal & "'"

    CrRelEstatisca.Formulas(10) = "CpPgtodepCapa   = '" & CpPagDepQtdeCapMal & "'"
    CrRelEstatisca.Formulas(11) = "CpPgtodepPgto   = '" & CpPagDepQtdePagMal & "'"
    CrRelEstatisca.Formulas(12) = "CpPgtodepChPgto = '" & CpPagDepQtdeChPagMal & "'"
    CrRelEstatisca.Formulas(13) = "CpDepPgtoDep    = '" & CpPagDepQtdeDepMal & "'"
    CrRelEstatisca.Formulas(14) = "CpdepPgtoChDep  = '" & CpPagDepQtdeChDepMal & "'"

    CrRelEstatisca.Formulas(15) = "CpPgtoCapaE    = '" & CppagQtdeCapEnv & "'"
    CrRelEstatisca.Formulas(16) = "CpPgtoPgtoE    = '" & CppagQtdePagEnv & "'"
    CrRelEstatisca.Formulas(17) = "CpPgtoChPgtoE = '" & CppagQtdeChPagEnv & "'"
    CrRelEstatisca.Formulas(18) = "CpPgtoDepE     = '" & CppagQtdeDepEnv & "'"
    CrRelEstatisca.Formulas(19) = "CpPgtoChDepE   = '" & CppagQtdeChDepEnv & "'"

    CrRelEstatisca.Formulas(20) = "CpDepCapaE     = '" & CpDepQtdeCapEnv & "'"
    CrRelEstatisca.Formulas(21) = "CpDepPgtoE     = '" & CpDepQtdePagEnv & "'"
    CrRelEstatisca.Formulas(22) = "CpDepChPgtoE   = '" & CpDepQtdeChPagEnv & "'"
    CrRelEstatisca.Formulas(23) = "CpDepDepE      = '" & CpDepQtdeDepEnv & "'"
    CrRelEstatisca.Formulas(24) = "CpDepChDepE    = '" & CpDepQtdeChDepEnv & "'"

    CrRelEstatisca.Formulas(25) = "CpPgtodepCapaE   = '" & CpPagDepQtdeCapEnv & "'"
    CrRelEstatisca.Formulas(26) = "CpPgtodepPgtoE   = '" & CpPagDepQtdePagEnv & "'"
    CrRelEstatisca.Formulas(27) = "CpPgtodepChPgtoE = '" & CpPagDepQtdeChPagEnv & "'"
    CrRelEstatisca.Formulas(28) = "CpDepPgtoDepE    = '" & CpPagDepQtdeDepEnv & "'"
    CrRelEstatisca.Formulas(29) = "CpdepPgtoChDepE  = '" & CpPagDepQtdeChDepEnv & "'"

    CrRelEstatisca.Formulas(30) = "DataProcessamento  = '" & DataFormatada & "'"
                                                            
    CrRelEstatisca.WindowState = crptMaximized
    CrRelEstatisca.WindowTitle = "Relat�rio de Estat�stica de Doctos no Caixa Expresso e Malote Empresa"
    CrRelEstatisca.Destination = crptToWindow
    CrRelEstatisca.Action = 1
    
    CrRelEstatisca.ReportFileName = Empty

End Sub
