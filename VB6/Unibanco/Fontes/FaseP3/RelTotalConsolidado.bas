Attribute VB_Name = "RelTotalConsolidado"
Option Explicit

Private qryGetResumoQuantidades As rdoQuery            ' query Contador de Documentos
Private RsAux                   As rdoResultset        ' Recordset

'* 1 - Capas s� com Pagamentos
Dim CppagQtdeCapMal        As Long      'Capas s� Com Pagamento Qtde Malotes
Dim CppagQtdeCapEnv        As Long      'Capas s� Com Pagamento Qtde Envelope
Dim CppagQtdePagMal        As Long      'Capas s� Com Pagamento Qtde Pagto Malote
Dim CppagQtdePagEnv        As Long      'Capas s� Com Pagamento Qtde Pagto Envelope
Dim CppagQtdeChPagMal      As Long      'Capas s� Com Pagamento Qtde Ch Pagto Malote
Dim CppagQtdeChPagEnv      As Long      'Capas s� Com Pagamento Qtde Ch Pagto Envelope
Dim CppagQtdeDepMal        As String    'Capas s� Com Pagamento Qtde Deposito Malote
Dim CppagQtdeDepEnv        As String    'Capas s� Com Pagamento Qtde Deposito Envelope
Dim CppagQtdeChDepMal      As String    'Capas s� Com Pagamento Qtde ch Deposito Malote
Dim CppagQtdeChDepEnv      As String    'Capas s� Com Pagamento Qtde Ch Deposito Envelope
Dim CpPagQtdeLIMal         As Long      'Capas s� com pagamento Qtde Lancto Interno Malote

'* 2 - Capas s� com Dep�sitos
Dim CpDepQtdeCapMal        As Long      'Capas s� com Dep�sitos Qtde Malote
Dim CpDepQtdeCapEnv        As Long      'Capas s� com Dep�sitos Qtde Envelope
Dim CpDepQtdePagMal        As String    'Capas s� com Dep�sitos Qtde Pagto Malote
Dim CpDepQtdePagEnv        As String    'Capas s� com Dep�sitos Qtde Pagto Envelope
Dim CpDepQtdeChPagMal      As String    'Capas s� com Dep�sitos Qtde Ch Pagto Malote
Dim CpDepQtdeChPagEnv      As String    'Capas s� com Dep�sitos Qtde Ch Pagto Envelope
Dim CpDepQtdeDepMal        As Long      'Capas s� com Dep�sitos Qtde Dep�sito Malote
Dim CpDepQtdeDepEnv        As Long      'Capas s� com Dep�sitos Qtde Dep�sito Envelope
Dim CpDepQtdeChDepMal      As Long      'Capas s� com Dep�sitos Qtde Ch Dep�sito Malote
Dim CpDepQtdeChDepEnv      As Long      'Capas s� com Dep�sitos Qtde Ch Dep�sito Envelope
Dim CpDepQtdeLIMal         As Long      'Capas s� com Dep�sitos Qtde Lancto Interno Malote

'* 3 - Capas Com Pagto/Deposito (Misto)
Dim CpPagDepQtdeCapMal     As Long 'Capas com Pagto/Dep�sito Qtde Malote
Dim CpPagDepQtdeCapEnv     As Long 'Capas com Pagto/Dep�sito Qtde Envelope
Dim CpPagDepQtdePagMal     As Long 'Capas com Pagto/Dep�sito Qtde Pagto Malote
Dim CpPagDepQtdePagEnv     As Long 'Capas com Pagto/Dep�sito Qtde Pagto Envelope
Dim CpPagDepQtdeChPagMal   As Long 'Capas com Pagto/Dep�sito Qtde Ch Pagto Malote
Dim CpPagDepQtdeChPagEnv   As Long 'Capas com Pagto/Dep�sito Qtde Ch Pagto Envelope
Dim CpPagDepQtdeDepMal     As Long 'Capas com Pagto/Dep�sito Qtde Dep�sito Malote
Dim CpPagDepQtdeDepEnv     As Long 'Capas com Pagto/Dep�sito Qtde Dep�sito Envelope
Dim CpPagDepQtdeChDepMal   As Long 'Capas com Pagto/Dep�sito Qtde Ch Dep�sito Malote
Dim CpPagDepQtdeChDepEnv   As Long 'Capas com Pagto/Dep�sito Qtde Ch Dep�sito Envelope
Dim CpPagDepQtdeLIMal      As Long 'Capas com Pagto/Dep�sito Qtde Lancto Interno Malote

'* 5 - Controlador de Ajustes
Dim CpAjPagtoE              As Long 'Capa com Pagto - Ajuste de D�bito / Cr�dito
Dim CpAjPagtoM              As Long 'Capa com Pagto - Ajuste de D�bito / Cr�dito
Dim CpAjDeptoE              As Long 'Capa com Dep�sito - Ajuste de D�bito / Cr�dito
Dim CpAjDeptoM              As Long 'Capa com Dep�sito - Ajuste de D�bito / Cr�dito
Dim CpAjPagDepE             As Long 'Capa com Pagto/Dep�sito - Ajuste de D�bito / Cr�dito
Dim CpAjPagDepM             As Long 'Capa com Pagto/Dep�sito - Ajuste de D�bito / Cr�dito

Private Sub Atualiza_Valores()

'* Valor default para  as vari�veis(Zero) *'

' * 1 - Capas s� com Pagamentos
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
    CpPagQtdeLIMal = 0          'Capas s� Com Pagamento Qtde LI Malote

' * 2 - Capas s� com Dep�sitos
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
    CpDepQtdeLIMal = 0          'Capas s� com Dep�sitos Qtde LI Malote
    
' * 3 - Capas Com Pagto/Deposito (Misto)
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
    CpPagDepQtdeLIMal = 0       'Capas com Pagto/Dep�sito Qtde LI Malote
    
' * 4 - Controle de Ajustes
    CpAjPagtoE = 0              'Capa com Pagto - Ajuste de D�bito / Cr�dito
    CpAjPagtoM = 0              'Capa com Pagto - Ajuste de D�bito / Cr�dito
    CpAjDeptoE = 0              'Capa com Dep�sito - Ajuste de D�bito / Cr�dito
    CpAjDeptoM = 0              'Capa com Dep�sito - Ajuste de D�bito / Cr�dito
    CpAjPagDepE = 0             'Capa com Pagto/Dep�sito - Ajuste de D�bito / Cr�dito
    CpAjPagDepM = 0             'Capa com Pagto/Dep�sito - Ajuste de D�bito / Cr�dito

    'LimpaFormulas
    Call LimpaFormulas

End Sub
Public Function RelEstatisticaConsolidado() As Boolean

Dim i           As Integer
Dim QryTimeOut  As Variant
Dim iCountList  As Integer, iTotalList As Integer

On Error GoTo Err_RelEstatisticaConsolidado

    'Aumenta timeout devido ao processamento demorado da Procedure
    QryTimeOut = Geral.Banco.QueryTimeout
    Geral.Banco.QueryTimeout = 300
    
    RelEstatisticaConsolidado = False
    
    Call Atualiza_Valores
    
    For i = 0 To frmRelTotalConsolidado.lstDias.ListCount - 1
        If frmRelTotalConsolidado.lstDias.Selected(i) Then
            iTotalList = iTotalList + 1
        End If
    Next
    
    If frmRelTotalConsolidado.lstDias.SelCount > 0 Then
    
        'Progess Bar
        iCountList = 0
        frmRelTotalConsolidado.pgbProcesso.Value = 0
        frmRelTotalConsolidado.pgbProcesso.Min = 0
        frmRelTotalConsolidado.pgbProcesso.Max = iTotalList
        frmRelTotalConsolidado.pgbProcesso.Visible = True
    
        For i = 0 To frmRelTotalConsolidado.lstDias.ListCount - 1
            
            If frmRelTotalConsolidado.lstDias.Selected(i) Then
                'Progress Bar
                iCountList = iCountList + 1
                frmRelTotalConsolidado.pgbProcesso.Value = iCountList - 0.5
                
                Set qryGetResumoQuantidades = Geral.Banco.CreateQuery("", "{call GetResumoQuantidades (?)}")
            
                Set RsAux = Nothing
                
                Screen.MousePointer = vbHourglass
        
                With qryGetResumoQuantidades
                    .rdoParameters(0).Value = frmRelTotalConsolidado.lstDias.ItemData(i) 'Data de movimento
                    .QueryTimeout = 300
                    Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
                End With
            
                'Progress Bar
                frmRelTotalConsolidado.pgbProcesso.Value = iCountList
                
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
                
            End If
        Next
    End If

    Screen.MousePointer = vbDefault
    frmRelTotalConsolidado.pgbProcesso.Visible = False
    
    If frmRelTotalConsolidado.fraGravar.Visible = False Then
        Call PreparaRelEstatistica
        Call Atualiza_Valores
    Else
        If Not GravaRelEstatistica Then GoTo Exit_RelEstatisticaConsolidado
    End If
    
    RelEstatisticaConsolidado = True
    
Exit_RelEstatisticaConsolidado:
    'Retorna timeout
    Geral.Banco.QueryTimeout = QryTimeOut
    Screen.MousePointer = vbDefault
    
    If Not (RsAux Is Nothing) Then Set RsAux = Nothing
    If Not (qryGetResumoQuantidades Is Nothing) Then Set qryGetResumoQuantidades = Nothing
    Exit Function

Err_RelEstatisticaConsolidado:
    Beep
    MsgBox "N�o foi poss�vel gerar o relat�rio, tente novamente", vbCritical + vbOKOnly, App.Title
    GoTo Exit_RelEstatisticaConsolidado
    
End Function
Private Sub CpPagtoQtdeCapa()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdeCapEnv = CppagQtdeCapEnv + RsAux!Qtde
        Else
           CppagQtdeCapMal = CppagQtdeCapMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoQtdePagto()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdePagEnv = CppagQtdePagEnv + RsAux!Qtde
        Else
           CppagQtdePagMal = CppagQtdePagMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults
    
End Sub
Private Sub CpPagtoQtdeChPagto()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CppagQtdeChPagEnv = CppagQtdeChPagEnv + RsAux!Qtde
        Else
           CppagQtdeChPagMal = CppagQtdeChPagMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepQtdeCapa()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeCapEnv = CpDepQtdeCapEnv + RsAux!Qtde
        Else
           CpDepQtdeCapMal = CpDepQtdeCapMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepQtdeDep()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeDepEnv = CpDepQtdeDepEnv + RsAux!Qtde
        Else
           CpDepQtdeDepMal = CpDepQtdeDepMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepQtdeChDep()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpDepQtdeChDepEnv = CpDepQtdeChDepEnv + RsAux!Qtde
        Else
           CpDepQtdeChDepMal = CpDepQtdeChDepMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoDepQtdeCapa()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeCapEnv = CpPagDepQtdeCapEnv + RsAux!Qtde
        Else
           CpPagDepQtdeCapMal = CpPagDepQtdeCapMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoDepQtdePagto()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdePagEnv = CpPagDepQtdePagEnv + RsAux!Qtde
        Else
           CpPagDepQtdePagMal = CpPagDepQtdePagMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpPagtoDepQtdeChPagto()

    While Not RsAux.EOF
        
        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeChPagEnv = CpPagDepQtdeChPagEnv + RsAux!Qtde
        Else
           CpPagDepQtdeChPagMal = CpPagDepQtdeChPagMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults
    
End Sub

Private Sub CpPagtoDepQtdeDep()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeDepEnv = CpPagDepQtdeDepEnv + RsAux!Qtde
        Else
           CpPagDepQtdeDepMal = CpPagDepQtdeDepMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend

    RsAux.MoreResults
    
End Sub
Private Sub CpPagtoDepQtdeChDep()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpPagDepQtdeChDepEnv = CpPagDepQtdeChDepEnv + RsAux!Qtde
        Else
           CpPagDepQtdeChDepMal = CpPagDepQtdeChDepMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
    
End Sub
Private Sub CpPagtoAjuste()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpAjPagtoE = CpAjPagtoE + RsAux!Qtde
        Else
           CpAjPagtoM = CpAjPagtoM + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
        
End Sub
Private Sub CpDeptoAjuste()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpAjDeptoE = CpAjDeptoE + RsAux!Qtde
        Else
           CpAjDeptoM = CpAjDeptoM + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
    
End Sub
Private Sub CpMistoAjuste()

    While Not RsAux.EOF

        If RsAux!IdEnv_Mal = "E" Then
           CpAjPagDepE = CpAjPagDepE + RsAux!Qtde
        Else
           CpAjPagDepM = CpAjPagDepM + RsAux!Qtde
        End If
    
        RsAux.MoveNext
        
    Wend
    
    RsAux.MoreResults
    
End Sub
Private Sub PreparaRelEstatistica()
'* Passagem de Parametros para Crystal Report *'

Dim DataFormatada As String, i As Integer

    Screen.MousePointer = vbHourglass
    
    DataFormatada = ""
    With frmRelTotalConsolidado.lstDias
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(DataFormatada) > 0 Then DataFormatada = DataFormatada & " - "
                DataFormatada = DataFormatada & Mid(.ItemData(i), 7, 2)
            End If
        Next
    End With
    
    Principal.RptGeral.ReportFileName = Empty

    Call LimpaFormulas

    With Principal.RptGeral
        .ReportFileName = App.path & "\RelTotalConsolidado.rpt "

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
        
        'Lancamento Interno
        .Formulas(39) = "CpPgtoLI         = '" & CpPagQtdeLIMal & "'"
        .Formulas(40) = "CpDepLI          = '" & CpDepQtdeLIMal & "'"
        .Formulas(41) = "CpDepPgtoLI      = '" & CpPagDepQtdeLIMal & "'"
        .Formulas(42) = "MesAnoMovimento  = '( " & frmRelTotalConsolidado.cmbMeses.List(frmRelTotalConsolidado.cmbMeses.ListIndex) & " )'"
        
        .WindowState = crptMaximized
        .WindowTitle = "Relat�rio de Estat�stica de Doctos no Caixa Expresso e Malote Empresa"
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
        For i = 0 To 42
            .Formulas(i) = Empty
        Next
    End With

End Sub
Private Sub CpPgtoLI()
    
    While Not RsAux.EOF
        If RsAux!IdEnv_Mal = "M" Then
           CpPagQtdeLIMal = CpPagQtdeLIMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepLI()
    
    While Not RsAux.EOF
        If RsAux!IdEnv_Mal = "M" Then
           CpDepQtdeLIMal = CpDepQtdeLIMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Sub CpDepPgtoLI()
    
    While Not RsAux.EOF
        If RsAux!IdEnv_Mal = "M" Then
           CpPagDepQtdeLIMal = CpPagDepQtdeLIMal + RsAux!Qtde
        End If
    
        RsAux.MoveNext
    Wend

    RsAux.MoreResults

End Sub
Private Function GravaRelEstatistica() As Boolean

Dim sArquivo As String, sLinha As String, sDelimit As String

On Error GoTo Err_GravaRelEstatistica

    GravaRelEstatistica = False
    sDelimit = ";"

    sArquivo = frmRelTotalConsolidado.txtDiretorio & "\" & Trim(frmRelTotalConsolidado.txtArquivo.Text) & ".csv"
     
    Open sArquivo For Binary Access Write As #1

    'Header Doctos Malote
    sLinha = "" & sDelimit
    sLinha = sLinha & "Qtd. Malote" & sDelimit
    sLinha = sLinha & "Qtd. Chq.Pagto." & sDelimit
    sLinha = sLinha & "Qtd. Contas" & sDelimit
    sLinha = sLinha & "Qtd. Deposito" & sDelimit
    sLinha = sLinha & "Qtd. Chq. Deposito" & sDelimit
    sLinha = sLinha & "Qtd. Lancto. Interno" & sDelimit
    sLinha = sLinha & "Qtd. Ajustes Deb/Cred" & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    '------------------------------------------
    '---            Malotes                 ---
    '------------------------------------------
    'Conte�do Doctos Malote s� com Pagtos
    sLinha = "Malote so com Pagtos" & sDelimit
    sLinha = sLinha & CppagQtdeCapMal & sDelimit
    sLinha = sLinha & CppagQtdeChPagMal & sDelimit
    sLinha = sLinha & CppagQtdePagMal & sDelimit
    sLinha = sLinha & CppagQtdeDepMal & sDelimit
    sLinha = sLinha & CppagQtdeChDepMal & sDelimit
    sLinha = sLinha & CpPagQtdeLIMal & sDelimit
    sLinha = sLinha & CpAjPagtoM & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    sLinha = "Malote so com Deposito" & sDelimit
    sLinha = sLinha & CpDepQtdeCapMal & sDelimit
    sLinha = sLinha & CpDepQtdeChPagMal & sDelimit
    sLinha = sLinha & CpDepQtdePagMal & sDelimit
    sLinha = sLinha & CpDepQtdeDepMal & sDelimit
    sLinha = sLinha & CpDepQtdeChDepMal & sDelimit
    sLinha = sLinha & CpDepQtdeLIMal & sDelimit
    sLinha = sLinha & CpAjDeptoM & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    sLinha = "Malote com Pagtos/Depositos" & sDelimit
    sLinha = sLinha & CpPagDepQtdeCapMal & sDelimit
    sLinha = sLinha & CpPagDepQtdeChPagMal & sDelimit
    sLinha = sLinha & CpPagDepQtdePagMal & sDelimit
    sLinha = sLinha & CpPagDepQtdeDepMal & sDelimit
    sLinha = sLinha & CpPagDepQtdeChDepMal & sDelimit
    sLinha = sLinha & CpPagDepQtdeLIMal & sDelimit
    sLinha = sLinha & CpAjPagDepM & vbCrLf
    'Grava registro
    Put #1, , sLinha
    '------------------------------------------
    '---            Envelopes               ---
    '------------------------------------------
    sLinha = "Envelope so com Pagtos" & sDelimit
    'Conte�do Doctos Envelopes s� com Pagtos
    sLinha = sLinha & CppagQtdeCapEnv & sDelimit
    sLinha = sLinha & CppagQtdeChPagEnv & sDelimit
    sLinha = sLinha & CppagQtdePagEnv & sDelimit
    sLinha = sLinha & CppagQtdeDepEnv & sDelimit
    sLinha = sLinha & CppagQtdeChDepEnv & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & CpAjDeptoE & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    sLinha = "Envelope so com Deposito" & sDelimit
    sLinha = sLinha & CpDepQtdeCapEnv & sDelimit
    sLinha = sLinha & CpDepQtdeChPagEnv & sDelimit
    sLinha = sLinha & CpDepQtdePagEnv & sDelimit
    sLinha = sLinha & CpDepQtdeDepEnv & sDelimit
    sLinha = sLinha & CpDepQtdeChDepEnv & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & CpAjPagtoE & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    sLinha = "Envelope com Pagtos/Depositos" & sDelimit
    sLinha = sLinha & CpPagDepQtdeCapEnv & sDelimit
    sLinha = sLinha & CpPagDepQtdeChPagEnv & sDelimit
    sLinha = sLinha & CpPagDepQtdePagEnv & sDelimit
    sLinha = sLinha & CpPagDepQtdeDepEnv & sDelimit
    sLinha = sLinha & CpPagDepQtdeChDepEnv & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & CpAjPagDepE & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    '------------------------------------------
    '---                TOTAL               ---
    '------------------------------------------
    'Conte�do Doctos Malote s� com Pagtos
    sLinha = "TOTAL Malote/Envelope so com Pagtos" & sDelimit
    sLinha = sLinha & (CppagQtdeCapMal + CppagQtdeCapEnv) & sDelimit
    sLinha = sLinha & (CppagQtdeChPagMal + CppagQtdeChPagEnv) & sDelimit
    sLinha = sLinha & (CppagQtdePagMal + CppagQtdePagEnv) & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & CpPagQtdeLIMal & sDelimit
    sLinha = sLinha & (CpAjPagtoM + CpAjPagtoE) & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    sLinha = "TOTAL Malote/Envelope so com Deposito" & sDelimit
    sLinha = sLinha & (CpDepQtdeCapMal + CpDepQtdeCapEnv) & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & "-" & sDelimit
    sLinha = sLinha & (CpDepQtdeDepMal + CpDepQtdeDepEnv) & sDelimit
    sLinha = sLinha & (CpDepQtdeChDepMal + CpDepQtdeChDepEnv) & sDelimit
    sLinha = sLinha & CpDepQtdeLIMal & sDelimit
    sLinha = sLinha & (CpAjDeptoM + CpAjDeptoE) & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    sLinha = "TOTAL Malote/Envelope com Pagtos/Depositos" & sDelimit
    sLinha = sLinha & (CpPagDepQtdeCapMal + CpPagDepQtdeCapEnv) & sDelimit
    sLinha = sLinha & (CpPagDepQtdeChPagMal + CpPagDepQtdeChPagEnv) & sDelimit
    sLinha = sLinha & (CpPagDepQtdePagMal + CpPagDepQtdePagEnv) & sDelimit
    sLinha = sLinha & (CpPagDepQtdeDepMal + CpPagDepQtdeDepEnv) & sDelimit
    sLinha = sLinha & (CpPagDepQtdeChDepMal + CpPagDepQtdeChDepEnv) & sDelimit
    sLinha = sLinha & CpPagDepQtdeLIMal & sDelimit
    sLinha = sLinha & (CpAjPagDepM + CpAjPagDepE) & vbCrLf
    'Grava registro
    Put #1, , sLinha
    
    '------------------------------------------
    '---           TOTAL GERAL              ---
    '------------------------------------------
    'Conte�do Doctos Malote s� com Pagtos
    sLinha = "TOTAL GERAL" & sDelimit
    
    sLinha = sLinha & (CppagQtdeCapMal + CppagQtdeCapEnv) + _
                      (CpDepQtdeCapMal + CpDepQtdeCapEnv) + _
                      (CpPagDepQtdeCapMal + CpPagDepQtdeCapEnv) & sDelimit
    
    sLinha = sLinha & (CppagQtdeChPagMal + CppagQtdeChPagEnv) + _
                      (CpPagDepQtdeChPagMal + CpPagDepQtdeChPagEnv) & sDelimit
    
    sLinha = sLinha & (CppagQtdePagMal + CppagQtdePagEnv) + _
                      (CpPagDepQtdePagMal + CpPagDepQtdePagEnv) & sDelimit
    
    sLinha = sLinha & (CpDepQtdeDepMal + CpDepQtdeDepEnv) + _
                      (CpPagDepQtdeDepMal + CpPagDepQtdeDepEnv) & sDelimit
    
    sLinha = sLinha & (CpDepQtdeChDepMal + CpDepQtdeChDepEnv) + _
                      (CpPagDepQtdeChDepMal + CpPagDepQtdeChDepEnv) & sDelimit
    
    sLinha = sLinha & CpPagQtdeLIMal + CpDepQtdeLIMal + CpPagDepQtdeLIMal & sDelimit
    
    sLinha = sLinha & (CpAjPagtoM + CpAjPagtoE) + _
                      (CpAjDeptoM + CpAjDeptoE) + _
                      (CpAjPagDepM + CpAjPagDepE) & vbCrLf
    'Grava registro
    Put #1, , sLinha
        
    'Fecha arquivo
    Close #1

    GravaRelEstatistica = True
    
    Exit Function
    
Err_GravaRelEstatistica:
    
    Beep
    Close #1
    MsgBox "N�o foi poss�vel gravar o arquivo." & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title

End Function
