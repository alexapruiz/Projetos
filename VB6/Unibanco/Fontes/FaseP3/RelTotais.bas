Attribute VB_Name = "RelTotais"
Private qryGetResumoQuantidades As rdoQuery            ' query Contador de Documentos
Private RsAux                   As rdoResultset        ' Recordset

'* 1 - Capas só com Pagamentos
Dim CppagQtdeCapMal        As Integer 'Capas só Com Pagamento Qtde Malotes
Dim CppagQtdeCapEnv        As Integer 'Capas só Com Pagamento Qtde Envelope
Dim CppagQtdePagMal        As Integer 'Capas só Com Pagamento Qtde Pagto Malote
Dim CppagQtdePagEnv        As Integer 'Capas só Com Pagamento Qtde Pagto Envelope
Dim CppagQtdeChPagMal      As Integer 'Capas só Com Pagamento Qtde Ch Pagto Malote
Dim CppagQtdeChPagEnv      As Integer 'Capas só Com Pagamento Qtde Ch Pagto Envelope
Dim CppagQtdeDepMal        As String  'Capas só Com Pagamento Qtde Deposito Malote
Dim CppagQtdeDepEnv        As String  'Capas só Com Pagamento Qtde Deposito Envelope
Dim CppagQtdeChDepMal      As String  'Capas só Com Pagamento Qtde ch Deposito Malote
Dim CppagQtdeChDepEnv      As String  'Capas só Com Pagamento Qtde Ch Deposito Envelope
Dim CpPagQtdeLIMal         As String  'Capas só com pagamento Qtde Lancto Interno Malote

'* 2 - Capas só com Depósitos
Dim CpDepQtdeCapMal        As Integer 'Capas só com Depósitos Qtde Malote
Dim CpDepQtdeCapEnv        As Integer 'Capas só com Depósitos Qtde Envelope
Dim CpDepQtdePagMal        As String  'Capas só com Depósitos Qtde Pagto Malote
Dim CpDepQtdePagEnv        As String  'Capas só com Depósitos Qtde Pagto Envelope
Dim CpDepQtdeChPagMal      As String  'Capas só com Depósitos Qtde Ch Pagto Malote
Dim CpDepQtdeChPagEnv      As String  'Capas só com Depósitos Qtde Ch Pagto Envelope
Dim CpDepQtdeDepMal        As Integer 'Capas só com Depósitos Qtde Depósito Malote
Dim CpDepQtdeDepEnv        As Integer 'Capas só com Depósitos Qtde Depósito Envelope
Dim CpDepQtdeChDepMal      As Integer 'Capas só com Depósitos Qtde Ch Depósito Malote
Dim CpDepQtdeChDepEnv      As Integer 'Capas só com Depósitos Qtde Ch Depósito Envelope
Dim CpDepQtdeLIMal         As String  'Capas só com Depósitos Qtde Lancto Interno Malote


'* 3 - Capas Com Pagto/Deposito (Misto)
Dim CpPagDepQtdeCapMal     As Integer 'Capas com Pagto/Depósito Qtde Malote
Dim CpPagDepQtdeCapEnv     As Integer 'Capas com Pagto/Depósito Qtde Envelope
Dim CpPagDepQtdePagMal     As Integer 'Capas com Pagto/Depósito Qtde Pagto Malote
Dim CpPagDepQtdePagEnv     As Integer 'Capas com Pagto/Depósito Qtde Pagto Envelope
Dim CpPagDepQtdeChPagMal   As Integer 'Capas com Pagto/Depósito Qtde Ch Pagto Malote
Dim CpPagDepQtdeChPagEnv   As Integer 'Capas com Pagto/Depósito Qtde Ch Pagto Envelope
Dim CpPagDepQtdeDepMal     As Integer 'Capas com Pagto/Depósito Qtde Depósito Malote
Dim CpPagDepQtdeDepEnv     As Integer 'Capas com Pagto/Depósito Qtde Depósito Envelope
Dim CpPagDepQtdeChDepMal   As Integer 'Capas com Pagto/Depósito Qtde Ch Depósito Malote
Dim CpPagDepQtdeChDepEnv   As Integer 'Capas com Pagto/Depósito Qtde Ch Depósito Envelope
Dim CpPagDepQtdeLIMal      As String  'Capas com Pagto/Depósito Qtde Lancto Interno Malote


'* 5 - Controlador de Ajustes
Dim CpAjPagtoE              As Integer 'Capa com Pagto - Ajuste de Débito / Crédito
Dim CpAjPagtoM              As Integer 'Capa com Pagto - Ajuste de Débito / Crédito
Dim CpAjDeptoE              As Integer 'Capa com Depósito - Ajuste de Débito / Crédito
Dim CpAjDeptoM              As Integer 'Capa com Depósito - Ajuste de Débito / Crédito
Dim CpAjPagDepE             As Integer 'Capa com Pagto/Depósito - Ajuste de Débito / Crédito
Dim CpAjPagDepM             As Integer 'Capa com Pagto/Depósito - Ajuste de Débito / Crédito

'* 6 - Última Transmissão Robo
Dim UltTransRobo0           As String  'Ultimo processamento Caixa Robo
Dim UltTransRobo1           As String  'Ultimo processamento Caixa Robo
Dim UltTransRobo2           As String  'Ultimo processamento Caixa Robo
Dim UltTransRobo3           As String  'Ultimo processamento Caixa Robo
Dim UltTransRobo4           As String  'Ultimo processamento Caixa Robo

Dim HrUltTransRobo0         As String  'Ultimo processamento Caixa Robo
Dim HrUltTransRobo1         As String  'Ultimo processamento Caixa Robo
Dim HrUltTransRobo2         As String  'Ultimo processamento Caixa Robo
Dim HrUltTransRobo3         As String  'Ultimo processamento Caixa Robo
Dim HrUltTransRobo4         As String  'Ultimo processamento Caixa Robo

Dim AgTransRobo0            As String  'Ultimo processamento Caixa Robo
Dim AgTransRobo1            As String  'Ultimo processamento Caixa Robo
Dim AgTransRobo2            As String  'Ultimo processamento Caixa Robo
Dim AgTransRobo3            As String  'Ultimo processamento Caixa Robo
Dim AgTransRobo4            As String  'Ultimo processamento Caixa Robo
Private Sub Atualiza_Valores()

'* Valor default para  as variáveis(Zero) *'

    '* 1 - Capas só com Pagamentos
    '* 2 - Capas só com Depósitos
    '* 3 - Capas Com Pagto/Deposito (Misto)
    '* 4 - Controle de Ajustes
    '* 5 - Última Transmissão Robo

' * 1 - Capas só com Pagamentos
    CppagQtdeCapMal = 0         'Capas só Com Pagamento Qtde Malotes
    CppagQtdeCapEnv = 0         'Capas só Com Pagamento Qtde Envelope
    CppagQtdePagMal = 0         'Capas só Com Pagamento Qtde Pagto Malote
    CppagQtdePagEnv = 0         'Capas só Com Pagamento Qtde Pagto Envelope
    CppagQtdeChPagMal = 0       'Capas só Com Pagamento Qtde Ch Pagto Malote
    CppagQtdeChPagEnv = 0       'Capas só Com Pagamento Qtde Ch Pagto Envelope
    CppagQtdeDepMal = "-"       'Capas só Com Pagamento Qtde Deposito Malote
    CppagQtdeDepEnv = "-"       'Capas só Com Pagamento Qtde Deposito Envelope
    CppagQtdeChDepMal = "-"     'Capas só Com Pagamento Qtde ch Deposito Malote
    CppagQtdeChDepEnv = "-"     'Capas só Com Pagamento Qtde Ch Deposito Envelope
    CpPagQtdeLIMal = 0          'Capas só Com Pagamento Qtde LI Malote

' * 2 - Capas só com Depósitos
    CpDepQtdeCapMal = 0         'Capas só com Depósitos Qtde Malote
    CpDepQtdeCapEnv = 0         'Capas só com Depósitos Qtde Envelope
    CpDepQtdePagMal = "-"       'Capas só com Depósitos Qtde Pagto Malote
    CpDepQtdePagEnv = "-"       'Capas só com Depósitos Qtde Pagto Envelope
    CpDepQtdeChPagMal = "-"     'Capas só com Depósitos Qtde Ch Pagto Malote
    CpDepQtdeChPagEnv = "-"     'Capas só com Depósitos Qtde Ch Pagto Envelope
    CpDepQtdeDepMal = 0         'Capas só com Depósitos Qtde Depósito Malote
    CpDepQtdeDepEnv = 0         'Capas só com Depósitos Qtde Depósito Envelope
    CpDepQtdeChDepMal = 0       'Capas só com Depósitos Qtde Ch Depósito Malote
    CpDepQtdeChDepEnv = 0       'Capas só com Depósitos Qtde Ch Depósito Envelope
    CpDepQtdeLIMal = 0          'Capas só com Depósitos Qtde LI Malote
    
' * 3 - Capas Com Pagto/Deposito (Misto)
    CpPagDepQtdeCapMal = 0      'Capas com Pagto/Depósito Qtde Malote
    CpPagDepQtdeCapEnv = 0      'Capas com Pagto/Depósito Qtde Envelope
    CpPagDepQtdePagMal = 0      'Capas com Pagto/Depósito Qtde Pagto Malote
    CpPagDepQtdePagEnv = 0      'Capas com Pagto/Depósito Qtde Pagto Envelope
    CpPagDepQtdeChPagMal = 0    'Capas com Pagto/Depósito Qtde Ch Pagto Malote
    CpPagDepQtdeChPagEnv = 0    'Capas com Pagto/Depósito Qtde Ch Pagto Envelope
    CpPagDepQtdeDepMal = 0      'Capas com Pagto/Depósito Qtde Depósito Malote
    CpPagDepQtdeDepEnv = 0      'Capas com Pagto/Depósito Qtde Depósito Envelope
    CpPagDepQtdeChDepMal = 0    'Capas com Pagto/Depósito Qtde Ch Depósito Malote
    CpPagDepQtdeChDepEnv = 0    'Capas com Pagto/Depósito Qtde Ch Depósito Envelope
    CpPagDepQtdeLIMal = 0       'Capas com Pagto/Depósito Qtde LI Malote
    
' * 4 - Controle de Ajustes
    CpAjPagtoE = 0              'Capa com Pagto - Ajuste de Débito / Crédito
    CpAjPagtoM = 0              'Capa com Pagto - Ajuste de Débito / Crédito
    CpAjDeptoE = 0              'Capa com Depósito - Ajuste de Débito / Crédito
    CpAjDeptoM = 0              'Capa com Depósito - Ajuste de Débito / Crédito
    CpAjPagDepE = 0             'Capa com Pagto/Depósito - Ajuste de Débito / Crédito
    CpAjPagDepM = 0             'Capa com Pagto/Depósito - Ajuste de Débito / Crédito

' * 5 - Última Transmissão Robo
    UltTransRobo0 = "-"          'Última Transmissão Robo
    UltTransRobo1 = "-"          'Última Transmissão Robo
    UltTransRobo2 = "-"          'Última Transmissão Robo
    UltTransRobo3 = "-"          'Última Transmissão Robo
    UltTransRobo4 = "-"          'Última Transmissão Robo
    UltTransRobo5 = "-"          'Última Transmissão Robo

    HrUltTransRobo0 = "-"        'Última Transmissão Robo
    HrUltTransRobo1 = "-"        'Última Transmissão Robo
    HrUltTransRobo2 = "-"        'Última Transmissão Robo
    HrUltTransRobo3 = "-"        'Última Transmissão Robo
    HrUltTransRobo4 = "-"        'Última Transmissão Robo
    HrUltTransRobo5 = "-"        'Última Transmissão Robo
    
    AgTransRobo0 = "-"           'Última Transmissão Robo
    AgTransRobo1 = "-"           'Última Transmissão Robo
    AgTransRobo2 = "-"           'Última Transmissão Robo
    AgTransRobo3 = "-"           'Última Transmissão Robo
    AgTransRobo4 = "-"           'Última Transmissão Robo
    AgTransRobo5 = "-"           'Última Transmissão Robo

    'LimpaFormulas
    Call LimpaFormulas

End Sub
Public Sub RelEstatistica()

Set qryGetResumoQuantidades = Geral.Banco.CreateQuery("", "{call GetResumoQuantidades (?)}")

    Set RsAux = Nothing
    
    Screen.MousePointer = vbHourglass

    With qryGetResumoQuantidades
        .rdoParameters(0).Value = Geral.DataProcessamento
        .QueryTimeout = 300
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
    
    UltTransRobo0 = "-"
    UltTransRobo1 = "-"
    UltTransRobo2 = "-"
    UltTransRobo3 = "-"
    UltTransRobo4 = "-"
    UltTransRobo5 = "-"

    HrUltTransRobo0 = "-"
    HrUltTransRobo1 = "-"
    HrUltTransRobo2 = "-"
    HrUltTransRobo3 = "-"
    HrUltTransRobo4 = "-"
    HrUltTransRobo5 = "-"
    
    AgTransRobo0 = "-"
    AgTransRobo1 = "-"
    AgTransRobo2 = "-"
    AgTransRobo3 = "-"
    AgTransRobo4 = "-"
    AgTransRobo5 = "-"
    
    '/Valores Default

    Call CpPagtoQtdeCapa
    Call CpPagtoQtdePagto
    Call CpPagtoQtdeChPagto
    Call CpPagtoAjuste
    Call CpPgtoLI
    
    Call CpDepQtdeCapa
    Call CpDepQtdeDep
    Call CpDepQtdeChDep
    Call CpDeptoAjuste
    Call CpDepLI
    
    Call CpPagtoDepQtdeCapa
    Call CpPagtoDepQtdePagto
    Call CpPagtoDepQtdeChPagto
    Call CpPagtoDepQtdeDep
    Call CpPagtoDepQtdeChDep
    Call CpDepPgtoLI
    Call CpMistoAjuste
    Call UltTransRobo
    
    Screen.MousePointer = vbDefault
    
    Call PreparaRelEstatistica
    Call Atualiza_Valores
    
    Set RsAux = Nothing
    Set qryGetResumoQuantidades = Nothing
    
End Sub
Private Sub CpPagtoQtdeCapa()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdeCapEnv = RsAux!Qtde
        Else
           CppagQtdeCapMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoQtdePagto()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdePagEnv = RsAux!Qtde
        Else
           CppagQtdePagMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults
    
End Sub
Private Sub CpPagtoQtdeChPagto()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdeChPagEnv = RsAux!Qtde
        Else
           CppagQtdeChPagMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepQtdeCapa()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeCapEnv = RsAux!Qtde
        Else
           CpDepQtdeCapMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepQtdeDep()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeDepEnv = RsAux!Qtde
        Else
           CpDepQtdeDepMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepQtdeChDep()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeChDepEnv = RsAux!Qtde
        Else
           CpDepQtdeChDepMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoDepQtdeCapa()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeCapEnv = RsAux!Qtde
        Else
           CpPagDepQtdeCapMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoDepQtdePagto()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdePagEnv = RsAux!Qtde
        Else
           CpPagDepQtdePagMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoDepQtdeChPagto()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeChPagEnv = RsAux!Qtde
        Else
           CpPagDepQtdeChPagMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults
    
End Sub

Private Sub CpPagtoDepQtdeDep()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeDepEnv = RsAux!Qtde
        Else
           CpPagDepQtdeDepMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults
    
End Sub
Private Sub CpPagtoDepQtdeChDep()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeChDepEnv = RsAux!Qtde
        Else
           CpPagDepQtdeChDepMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
    
End Sub
Private Sub CpPagtoAjuste()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpAjPagtoE = RsAux!Qtde
        Else
           CpAjPagtoM = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
        
End Sub
Private Sub CpDeptoAjuste()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpAjDeptoE = RsAux!Qtde
        Else
           CpAjDeptoM = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
    
End Sub
Private Sub CpMistoAjuste()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpAjPagDepE = RsAux!Qtde
        Else
           CpAjPagDepM = RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
    
End Sub
Private Sub UltTransRobo()

    If RsAux.EOF = False Then
        AgTransRobo0 = IIf(IsNull(RsAux!Caixa), AgTransRobo0, RsAux!Caixa)
        UltTransRobo0 = IIf(IsNull(RsAux!UltTransRobo), UltTransRobo0, RsAux!UltTransRobo)
        HrUltTransRobo0 = IIf(IsNull(RsAux!UltTransRobo), HrUltTransRobo0, RsAux!UltTransRobo)
        If Not IsNull(RsAux!UltTransRobo) Then
            UltTransRobo0 = Format(UltTransRobo0, "DD/MM/YYYY")
            HrUltTransRobo0 = Format(HrUltTransRobo0, "HH:MM:SS")
        End If
    Else
        Exit Sub
    End If
        
    RsAux.MoveNext
        
    If RsAux.EOF = False Then
        AgTransRobo1 = IIf(IsNull(RsAux!Caixa), AgTransRobo1, RsAux!Caixa)
        UltTransRobo1 = IIf(IsNull(RsAux!UltTransRobo), UltTransRobo1, RsAux!UltTransRobo)
        HrUltTransRobo1 = IIf(IsNull(RsAux!UltTransRobo), HrUltTransRobo1, RsAux!UltTransRobo)
        If Not IsNull(RsAux!UltTransRobo) Then
            UltTransRobo1 = Format(UltTransRobo1, "DD/MM/YYYY")
            HrUltTransRobo1 = Format(HrUltTransRobo1, "HH:MM:SS")
        End If
    Else
        Exit Sub
    End If
    
    RsAux.MoveNext
    
    If RsAux.EOF = False Then
        AgTransRobo2 = IIf(IsNull(RsAux!Caixa), AgTransRobo2, RsAux!Caixa)
        UltTransRobo2 = IIf(IsNull(RsAux!UltTransRobo), UltTransRobo2, RsAux!UltTransRobo)
        HrUltTransRobo2 = IIf(IsNull(RsAux!UltTransRobo), HrUltTransRobo2, RsAux!UltTransRobo)
        If Not IsNull(RsAux!UltTransRobo) Then
            UltTransRobo2 = Format(UltTransRobo2, "DD/MM/YYYY")
            HrUltTransRobo2 = Format(HrUltTransRobo2, "HH:MM:SS")
        End If
    Else
        Exit Sub
    End If

    
    RsAux.MoveNext
    
    If RsAux.EOF = False Then
        AgTransRobo3 = IIf(IsNull(RsAux!Caixa), AgTransRobo3, RsAux!Caixa)
        UltTransRobo3 = IIf(IsNull(RsAux!UltTransRobo), UltTransRobo3, RsAux!UltTransRobo)
        HrUltTransRobo3 = IIf(IsNull(RsAux!UltTransRobo), HrUltTransRobo3, RsAux!UltTransRobo)
        If Not IsNull(RsAux!UltTransRobo) Then
            UltTransRobo3 = Format(UltTransRobo3, "DD/MM/YYYY")
            HrUltTransRobo3 = Format(HrUltTransRobo3, "HH:MM:SS")
        End If
    Else
        Exit Sub
    End If
    
    RsAux.MoveNext
    
    If RsAux.EOF = False Then
        AgTransRobo4 = IIf(IsNull(RsAux!Caixa), AgTransRobo4, RsAux!Caixa)
        UltTransRobo4 = IIf(IsNull(RsAux!UltTransRobo), UltTransRobo4, RsAux!UltTransRobo)
        HrUltTransRobo4 = IIf(IsNull(RsAux!UltTransRobo), HrUltTransRobo4, RsAux!UltTransRobo)
        If Not IsNull(RsAux!UltTransRobo) Then
            UltTransRobo4 = Format(UltTransRobo4, "DD/MM/YYYY")
            HrUltTransRobo4 = Format(HrUltTransRobo4, "HH:MM:SS")
        End If
    Else
        Exit Sub
    End If
                
End Sub
Private Sub PreparaRelEstatistica()
'* Passagem de Parametros para Crystal Report *'

Dim DataFormatada As String

    Screen.MousePointer = vbHourglass

    DataFormatada = Mid(Geral.DataProcessamento, 7, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Mid(Geral.DataProcessamento, 1, 4)

    Principal.RptGeral.ReportFileName = Empty

    Call LimpaFormulas

    With Principal.RptGeral
        .ReportFileName = App.path & "\estatistica.rpt "

        .Formulas(0) = "CpPgtoCapa    = '" & CppagQtdeCapMal & "'"
        .Formulas(1) = "CpPgtoPgto    = '" & CppagQtdePagMal & "'"
        .Formulas(2) = "CpPgtoChPagto = '" & CppagQtdeChPagMal & "'"
        .Formulas(3) = "CpPgtoDep     = '" & CppagQtdeDepMal & "'"
        .Formulas(4) = "CpPgtoChDep   = '" & CppagQtdeChDepMal & "'"
        
        .Formulas(5) = "CpDepCapa     = '" & CpDepQtdeCapMal & "'"
        .Formulas(6) = "CpDepPgto     = '" & CpDepQtdePagMal & "'"
        .Formulas(7) = "CpDepChPgto   = '" & CpDepQtdeChPagMal & "'"
        .Formulas(8) = "CpDepDep      = '" & CpDepQtdeDepMal & "'"
        .Formulas(9) = "CpDepChDep    = '" & CpDepQtdeChDepMal & "'"
        
        .Formulas(10) = "CpPgtodepCapa   = '" & CpPagDepQtdeCapMal & "'"
        .Formulas(11) = "CpPgtodepPgto   = '" & CpPagDepQtdePagMal & "'"
        .Formulas(12) = "CpPgtodepChPgto = '" & CpPagDepQtdeChPagMal & "'"
        .Formulas(13) = "CpDepPgtoDep    = '" & CpPagDepQtdeDepMal & "'"
        .Formulas(14) = "CpdepPgtoChDep  = '" & CpPagDepQtdeChDepMal & "'"
        
        .Formulas(15) = "CpPgtoCapaE    = '" & CppagQtdeCapEnv & "'"
        .Formulas(16) = "CpPgtoPgtoE    = '" & CppagQtdePagEnv & "'"
        .Formulas(17) = "CpPgtoChPgtoE = '" & CppagQtdeChPagEnv & "'"
        .Formulas(18) = "CpPgtoDepE     = '" & CppagQtdeDepEnv & "'"
        .Formulas(19) = "CpPgtoChDepE   = '" & CppagQtdeChDepEnv & "'"
        
        .Formulas(20) = "CpDepCapaE     = '" & CpDepQtdeCapEnv & "'"
        .Formulas(21) = "CpDepPgtoE     = '" & CpDepQtdePagEnv & "'"
        .Formulas(22) = "CpDepChPgtoE   = '" & CpDepQtdeChPagEnv & "'"
        .Formulas(23) = "CpDepDepE      = '" & CpDepQtdeDepEnv & "'"
        .Formulas(24) = "CpDepChDepE    = '" & CpDepQtdeChDepEnv & "'"
        
        .Formulas(25) = "CpPgtodepCapaE   = '" & CpPagDepQtdeCapEnv & "'"
        .Formulas(26) = "CpPgtodepPgtoE   = '" & CpPagDepQtdePagEnv & "'"
        .Formulas(27) = "CpPgtodepChPgtoE = '" & CpPagDepQtdeChPagEnv & "'"
        .Formulas(28) = "CpDepPgtoDepE    = '" & CpPagDepQtdeDepEnv & "'"
        .Formulas(29) = "CpdepPgtoChDepE  = '" & CpPagDepQtdeChDepEnv & "'"
        .Formulas(30) = "DataProcessamento  = '" & DataFormatada & "'"

        .Formulas(31) = "CpAjPagtoM    = '" & CpAjPagtoM & "'"
        .Formulas(32) = "CpAjDepM      = '" & CpAjDeptoM & "'"
        .Formulas(33) = "CpAjPagtoDepM = '" & CpAjPagDepM & "'"
        .Formulas(34) = "CpAjPagE      = '" & CpAjPagtoE & "'"
        .Formulas(35) = "CpAjDepE      = '" & CpAjDeptoE & "'"
        .Formulas(36) = "CpAjPagDepE   = '" & CpAjPagDepE & "'"
        .Formulas(38) = "AgenciaCentral    = '" & Geral.AgenciaCentral & "'"
        
        .Formulas(39) = "AgTranRobo0   = '" & AgTransRobo0 & "'"
        .Formulas(40) = "AgTranRobo1   = '" & AgTransRobo1 & "'"
        .Formulas(41) = "AgTranRobo2   = '" & AgTransRobo2 & "'"
        .Formulas(42) = "AgTranRobo3   = '" & AgTransRobo3 & "'"
        .Formulas(43) = "AgTranRobo4   = '" & AgTransRobo4 & "'"
        
        .Formulas(44) = "UltTransRobo0  = '" & UltTransRobo0 & "'"
        .Formulas(45) = "UltTransRobo1  = '" & UltTransRobo1 & "'"
        .Formulas(46) = "UltTransRobo2  = '" & UltTransRobo2 & "'"
        .Formulas(47) = "UltTransRobo3  = '" & UltTransRobo3 & "'"
        .Formulas(48) = "UltTransRobo4  = '" & UltTransRobo4 & "'"
        
        .Formulas(49) = "HrUltTransRobo0  = '" & HrUltTransRobo0 & "'"
        .Formulas(50) = "HrUltTransRobo1  = '" & HrUltTransRobo1 & "'"
        .Formulas(51) = "HrUltTransRobo2  = '" & HrUltTransRobo2 & "'"
        .Formulas(52) = "HrUltTransRobo3  = '" & HrUltTransRobo3 & "'"
        .Formulas(53) = "HrUltTransRobo4  = '" & HrUltTransRobo4 & "'"
        
        'Lancamento Interno
        .Formulas(54) = "CpPgtoLI         = '" & CpPagQtdeLIMal & "'"
        .Formulas(55) = "CpDepLI          = '" & CpDepQtdeLIMal & "'"
        .Formulas(56) = "CpDepPgtoLI      = '" & CpPagDepQtdeLIMal & "'"
        
        .WindowState = crptMaximized
        .WindowTitle = "Relatório de Estatística de Doctos no Caixa Expresso e Malote Empresa"
        .Destination = crptToWindow
        .Action = 1
    End With
    
    Screen.MousePointer = vbDefault
    
    With Principal.RptGeral
        .ReportFileName = Empty
        .WindowState = Empty
        .WindowTitle = Empty
        .Destination = Empty
    End With
    
    Call LimpaFormulas

End Sub

Private Sub LimpaFormulas()

Dim i As Integer

    With Principal.RptGeral
        For i = 0 To 56
            .Formulas(i) = Empty
        Next
    End With

End Sub
Private Sub CpPgtoLI()
    
    CpPagQtdeLIMal = 0
    While Not RsAux.EOF
        If RsAux!IdEnv_Mal = "M" Then
           CpPagQtdeLIMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepLI()
    
    CpDepQtdeLIMal = 0
    While Not RsAux.EOF
        If RsAux!IdEnv_Mal = "M" Then
           CpDepQtdeLIMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepPgtoLI()
    
    CpPagDepQtdeLIMal = 0
    While Not RsAux.EOF
        If RsAux!IdEnv_Mal = "M" Then
           CpPagDepQtdeLIMal = RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
