VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalculoBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"DataDeposito"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''
'Definição de Variaveis'
''''''''''''''''''''''''
Private m_cDataDeposito         As DataDeposito
Private m_DataProcessamento     As Long
Private m_IdBordero             As Long
Private m_Connection            As New ADODB.Connection
Private m_BorderoDivergente     As Boolean
Private m_QtdMaxDatas           As Integer
Private m_QtdMaxCheques         As Long
Private m_Icheque               As Boolean

Public DataDeposito             As New DataDeposito
Public Property Get BorderoDivergente() As Boolean
    BorderoDivergente = m_BorderoDivergente
End Property
Public Function Calcula() As Boolean

    Dim i                   As Integer
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim rst                 As New ADODB.Recordset
    
    Calcula = False
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se não existe nenhuma data de depósito, seleciona todas'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DataDeposito.Count = 0 Then
        
        Set rst = m_Connection.Execute(Proc_Selecionar.GetDatasBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero))

        Do While Not rst.EOF()
            Me.DataDeposito.Add rst!DataDeposito
            Me.DataDeposito(rst.AbsolutePosition).DataProcessamento = m_DataProcessamento
            Me.DataDeposito(rst.AbsolutePosition).Quantidade = rst!QuantidadeCheques
            Me.DataDeposito(rst.AbsolutePosition).ValorDeposito = rst!ValorDeposito
            Me.DataDeposito(rst.AbsolutePosition).IdBordero = m_IdBordero
            rst.MoveNext
        Loop
        
        rst.Close

    End If
    
    ''''''''''''''''''''''''''''''''''''
    'Calcula todas as datas de deposito'
    ''''''''''''''''''''''''''''''''''''
    m_BorderoDivergente = True
    For i = 1 To Me.DataDeposito.Count
    
        If Not Me.DataDeposito(i).Calcula(m_Connection) Then
        
            m_BorderoDivergente = False
            Calcula = False
        End If
    Next i
    
    '''''''''''''''''''''''''''
    'Define bordero Divergente'
    '''''''''''''''''''''''''''
    If m_BorderoDivergente Then Calcula = True
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Calcula os cheques que estao fora do Valor do Deposito com base no Aviso de Diferenca'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CalculaChequesAD(ByVal pValorAD As Double) As Boolean

    Dim i                   As Integer
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim rst                 As New ADODB.Recordset
    
    CalculaChequesAD = False
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se não existe nenhuma data de depósito, seleciona todas'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DataDeposito.Count = 0 Then
        
        Set rst = m_Connection.Execute(Proc_Selecionar.GetDatasBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero))

        Do While Not rst.EOF()
            Me.DataDeposito.Add rst!DataDeposito
            Me.DataDeposito(rst.AbsolutePosition).DataProcessamento = m_DataProcessamento
            Me.DataDeposito(rst.AbsolutePosition).Quantidade = rst!QuantidadeCheques
            Me.DataDeposito(rst.AbsolutePosition).ValorDeposito = rst!ValorDeposito
            Me.DataDeposito(rst.AbsolutePosition).IdBordero = m_IdBordero
            rst.MoveNext
        Loop
        
        rst.Close

    End If
    
    ''''''''''''''''''''''''''''''''''''
    'Calcula todas as datas de deposito'
    ''''''''''''''''''''''''''''''''''''
    For i = Me.DataDeposito.Count To 1 Step -1
    
        Me.DataDeposito(i).ValorAD = pValorAD
        If Not Me.DataDeposito(i).CalculaChequesAD(m_Connection) Then
            CalculaChequesAD = False
            Exit For
        End If
        
    Next i
    
    
    
End Function

Public Function VoltaStatusChequesIndevidos()
    Dim qStr, rStr, sStr              As String
      
    On Error GoTo Erro_VoltaStatusChequesIndevidos:
    
   'Atualiza Status do(s) cheque(s) indevidos
    qStr = "UPDATE CHEQUE SET STATUS = '1'"
    qStr = qStr & " WHERE DataProcessamento = " & m_DataProcessamento
    qStr = qStr & " AND IdBordero = " & m_IdBordero
    qStr = qStr & " AND Status = 'I'"
    
   'Atualiza Status do(s) cheque(s) e da(s) DataDeposito indevidos
    rStr = "UPDATE DATADEPOSITO SET STATUS = '1'"
    rStr = rStr & " WHERE DataProcessamento = " & m_DataProcessamento
    rStr = rStr & " AND IdBordero = " & m_IdBordero
    rStr = rStr & " AND Status = 'I'"
        
    m_Connection.BeginTrans
    
    Call m_Connection.Execute(qStr)
    Call m_Connection.Execute(rStr)
        
   'Se não houve erro Atualiza propriedade de controle e executa Commit
    m_Connection.CommitTrans
    Me.VerificaChequeIndevidos = True
    Exit Function
    
Erro_VoltaStatusChequesIndevidos:
   'Se ouve erro Atualiza propriedade p/ false e executa RollBack
    m_Connection.RollbackTrans
    Me.VerificaChequeIndevidos = False
    Call TratamentoErro("Falha ao Voltar Status dos Cheques indevidos", Err)
    
End Function
Public Function CalculaChequesIndevidosQTDE(Optional ByRef pQtdChequesAlterados As Long, Optional ByVal pDevolverCheques As Boolean = False) As Boolean
       
    Dim sStr, pStr          As String
    Dim lRetorno            As Long
    Dim rstCountCheques     As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim Repete              As Boolean
    Dim IdDataDeposito      As Integer
    Dim lDataDeposito       As Long
    Dim rstDataDeposito     As New ADODB.Recordset
    Dim rstCheques          As New ADODB.Recordset
      
    On Error GoTo Erro_CalculaChequesIndevidosQTDE:
    
   'Verifica se foi executado retorno dos Status
    If Not Me.VerificaChequeIndevidos Then
        MsgBox "É preciso executar o método 'Retornar Status dos Cheques'."
        Exit Function
    End If
    
    Set rstCountCheques = m_Connection.Execute(Proc_Selecionar.GetChequesIndevidosQTDE(m_DataProcessamento, m_IdBordero))
                                               
    If Not rstCountCheques.EOF() Then
       'Verifica se há mais cheques que o permitido
        If rstCountCheques.RecordCount > m_QtdMaxCheques Then
           'Move ponteiro p/ 1o. cheque indevido
            rstCountCheques.Move (m_QtdMaxCheques)
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se não existe nenhuma data de depósito, seleciona todas'
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Me.DataDeposito.Count = 0 And pDevolverCheques = True Then
                
                Set rstDataDeposito = m_Connection.Execute(Proc_Selecionar.GetDatasBordero( _
                                               m_DataProcessamento, _
                                               m_IdBordero))

                Do While Not rstDataDeposito.EOF()
                    Me.DataDeposito.Add rstDataDeposito!DataDeposito, rstDataDeposito!DataDeposito
                    Me.DataDeposito(rstDataDeposito.AbsolutePosition).DataProcessamento = m_DataProcessamento
                    Me.DataDeposito(rstDataDeposito.AbsolutePosition).Quantidade = rstDataDeposito!QuantidadeCheques
                    Me.DataDeposito(rstDataDeposito.AbsolutePosition).ValorDeposito = rstDataDeposito!ValorDeposito
                    Me.DataDeposito(rstDataDeposito.AbsolutePosition).IdBordero = m_IdBordero

                    Set rstCheques = g_cMainConnection.Execute(Proc_Selecionar.GetCheques( _
                                                               m_DataProcessamento, _
                                                               m_IdBordero, _
                                                               rstDataDeposito!DataDeposito))
                    Do While Not rstCheques.EOF()
                        Me.DataDeposito(rstDataDeposito.AbsolutePosition).Cheque.Add rstCheques!CMC7, rstCheques!CMC7
                        rstCheques.MoveNext
                    Loop

                    rstCheques.Close
                    rstDataDeposito.MoveNext
                Loop
                rstDataDeposito.Close
            End If
            
            IdDataDeposito = 1
            If Not rstCountCheques.EOF Then
                lDataDeposito = rstCountCheques!DataDeposito
            End If
            
           'Adiciona idcheque a string p/ montar query de atualizacao
            While Not rstCountCheques.EOF
                If Len(pStr) > 0 Then
                    pStr = pStr & ","
                End If
                
                If rstCountCheques!DataDeposito <> lDataDeposito Then
                    lDataDeposito = rstCountCheques!DataDeposito
                    IdDataDeposito = IdDataDeposito + 1
                End If
                
                pStr = pStr & rstCountCheques(0).Value
                
                If pDevolverCheques = True Then
                    Me.DataDeposito(IdDataDeposito).Cheque.Item(rstCountCheques!CMC7).Status = "I"
                    Me.DataDeposito(IdDataDeposito).Cheque.Item(rstCountCheques!CMC7).IdCheque = rstCountCheques!IdCheque
                    Me.DataDeposito(IdDataDeposito).Cheque.Item(rstCountCheques!CMC7).Valor = CDbl(rstCountCheques!Valor)
                End If
                
                rstCountCheques.MoveNext
            Wend


           'Atualiza Status do(s) cheque(s) indevidos
            sStr = "UPDATE CHEQUE SET STATUS = 'I'"
            sStr = sStr & " WHERE DataProcessamento = " & m_DataProcessamento
            sStr = sStr & " AND IdBordero = " & m_IdBordero
            sStr = sStr & " AND Idcheque IN (" & pStr & ")"
            
AtualizaCheque:
            m_Connection.BeginTrans
            
            Call m_Connection.Execute(sStr, lRetorno, adCmdText)
            
            If lRetorno = 0 Then
                m_Connection.RollbackTrans
                Repete = True
                Err.Raise 996, App.Title, "houve Falha no Processo de Atualização de Cheques Indevidos(Qtde)"
            Else
                pQtdChequesAlterados = lRetorno
                m_Connection.CommitTrans
            End If
        End If
        
    End If
    
    CalculaChequesIndevidosQTDE = True
    
    Exit Function
    
Erro_CalculaChequesIndevidosQTDE:
    Call TratamentoErro("Houve Falha na Seleção/Atualização de Cheques indevidos", Err, Repete)
    
    If Repete Then
        Resume AtualizaCheque
        Repete = False
    End If
    
    CalculaChequesIndevidosQTDE = False

End Function
Public Function CalculaChequesIndevidosDATA(Optional ByRef pQtdDatasAlteradas As Long) As Boolean
       
    Dim sStr, pStr           As String
    Dim Proc_Selecionar      As New Custodia.Selecionar
    Dim rstCountDatas        As New ADODB.Recordset
    Dim lRetorno1, lRetorno2 As Long
    Dim Repete               As Boolean
    
   'Verifica se foi executado retorno dos Status
    On Error GoTo Erro_CalculaChequesIndevidosDATA:
    
    If Not Me.VerificaChequeIndevidos Then
        MsgBox "É preciso executar o método 'Retornar Status dos Cheques'."
        Exit Function
    End If
    
    Set rstCountDatas = m_Connection.Execute(Proc_Selecionar.GetChequesIndevidosDATA(m_DataProcessamento, m_IdBordero))
                                               
    If Not rstCountDatas.EOF() Then
       'Verifica se há mais cheques que o permitido
        If rstCountDatas.RecordCount > m_QtdMaxDatas Then
           'Move ponteiro p/ 1o. cheque indevido
            rstCountDatas.Move (m_QtdMaxDatas)
            
           'Adiciona idcheque a string p/ montar query de atualizacao
            While Not rstCountDatas.EOF
                If Len(pStr) > 0 Then
                    pStr = pStr & ","
                End If
                
                pStr = pStr & rstCountDatas(0).Value
                rstCountDatas.MoveNext

            Wend
            
           'Atualiza Status do(s) cheque(s) indevidos
            sStr = "UPDATE CHEQUE SET STATUS = 'I'"
            sStr = sStr & " WHERE DataProcessamento = " & m_DataProcessamento
            sStr = sStr & " AND IdBordero = " & m_IdBordero
            sStr = sStr & " AND DataDeposito IN (" & pStr & ")"
            
           'Atualiza Status da(s) Data(s) de depósito(s) indevidos
            rStr = "UPDATE DATADEPOSITO SET STATUS = 'I'"
            rStr = rStr & " WHERE DataProcessamento = " & m_DataProcessamento
            rStr = rStr & " AND IdBordero = " & m_IdBordero
            rStr = rStr & " AND DataDeposito IN (" & pStr & ")"
            
AtualizaCheque:
            m_Connection.BeginTrans

            Call m_Connection.Execute(rStr, lRetorno1, adCmdText)
            Call m_Connection.Execute(sStr, lRetorno2, adCmdText)
            
           'obs.: LRetorno2 pode ser 0, pois pode não ter sido incluido na Tab cheque todos os cheques indicados no borderô.
            If lRetorno1 = 0 Then
                Err.Raise 996, App.Title, "Houve Falha no Processo de Atualização de Cheques Indevidos.(Data)"
            Else
                m_Connection.CommitTrans
                pQtdDatasAlteradas = lRetorno1
            End If
        End If
    End If
    
    CalculaChequesIndevidosDATA = True
    
    Exit Function
            
Erro_CalculaChequesIndevidosDATA:
    m_Connection.RollbackTrans
    Call TratamentoErro("Houve Falha na Seleção/Atualização de Cheques indevidos", Err, Repete)
    
    If Repete Then
        Resume AtualizaCheque
        Repete = False
    End If
    
    CalculaChequesIndevidosDATA = False
End Function
Public Property Let IdBordero(ByVal vData As Long)
    m_IdBordero = vData
End Property
Public Property Get IdBordero() As Long
    IdBordero = m_IdBordero
End Property
Public Property Let DataProcessamento(ByVal vData As Long)
    m_DataProcessamento = vData
End Property
Public Property Get DataProcessamento() As Long
    DataProcessamento = m_DataProcessamento
End Property
Public Property Let QuantidadeMaximaCheques(ByVal pQtdMaxCheques As Long)
    m_QtdMaxCheques = pQtdMaxCheques
End Property
Public Property Let QuantidadeMaximaDatas(ByVal pQtdMaxDatas As Integer)
    m_QtdMaxDatas = pQtdMaxDatas
End Property
Public Sub SetConnection(ByVal pConnection As ADODB.Connection)
    Set m_Connection = pConnection
End Sub
Public Property Let VerificaChequeIndevidos(ByVal Icheque As Boolean)
    m_Icheque = Icheque
End Property
Public Property Get VerificaChequeIndevidos() As Boolean
    VerificaChequeIndevidos = m_Icheque
End Property
Private Sub Class_Initialize()
  'Obriga a execução do método (VoltaStatusChequesIndevidos), antes de verificar datas e qtde's
   Me.VerificaChequeIndevidos = False
End Sub
Private Sub Class_Terminate()
    Set m_cDataDeposito = Nothing
End Sub
