VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Excluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'''''''''''''''''''''''''''''
'Remove o Aviso de Diferenca'
'''''''''''''''''''''''''''''
    Public Function RemoveAvisoDiferenca(ByVal pDataProcessamento As Long, _
                                     ByVal pIdBordero As Long) As String
                                     
    Dim sStr        As String

           sStr = "DELETE FROM AVISODIFERENCA"
    sStr = sStr & " WHERE DataProcessamento < " & pDataProcessamento
    sStr = sStr & " AND IdBordero = " & pIdBordero

    RemoveAvisoDiferenca = sStr

End Function


Public Function RemoveChequesIndevidos(ByVal pDataProcessamento As Long, _
                                       ByVal pIdBordero As Long, _
                                       ByVal pDataDeposito As Long) As String

    Dim sStr        As String
    
           sStr = "DELETE"
    sStr = sStr & "  FROM CHEQUE"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    sStr = sStr & "   AND Status = 'I'"
    
    RemoveChequesIndevidos = sStr
End Function

Public Function RemoveChequesDeletados(ByVal pDataProcessamento As Long, _
                                       ByVal pIdBordero As Long, _
                                       ByVal pDataDeposito As Long) As String

    Dim sStr        As String
    
           sStr = "DELETE"
    sStr = sStr & "  FROM CHEQUE"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    'sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    sStr = sStr & "   AND Status = 'D'"
    
    RemoveChequesDeletados = sStr
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Exclui o(s) registros data tabela DataDeposito referente à Data de Deposito'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveDataDeposito(ByVal pDataProcessamento As Long, _
                                   ByVal pIdBordero As Long, _
                                   ByVal pDataDeposito As Long) As String

    Dim sStr        As String
    
           sStr = "DELETE"
    sStr = sStr & " FROM DataDeposito"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & " AND IdBordero = " & pIdBordero
    sStr = sStr & " AND DataDeposito = " & pDataDeposito
    
    RemoveDataDeposito = sStr

End Function
Public Function LimpaDataDeposito(ByVal pDataProcessamento As Long) As String

    Dim sStr        As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM DataDeposito"
    sStr = sStr & " WHERE DataProcessamento < " & pDataProcessamento
    
    LimpaDataDeposito = sStr

End Function
Public Function LimpaBordero(ByVal pDataProcessamento As Long) As String

    Dim sStr        As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM Bordero"
    sStr = sStr & " WHERE DataProcessamento < " & Trim(Str(pDataProcessamento))
    
    LimpaBordero = sStr

End Function
Public Function LimpaCheque(ByVal pDataProcessamento As Long) As String

    Dim sStr As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento < " & Trim(Str(pDataProcessamento))
    
    LimpaCheque = sStr

End Function
Public Function LimpaRejeicaoRemessa(ByVal pDataProcessamento As Long) As String

    Dim sStr        As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM RejeicaoRemessa"
    sStr = sStr & " WHERE DataProcessamento < " & Trim(Str(pDataProcessamento))
    
    LimpaRejeicaoRemessa = sStr

End Function
Public Function LimpaChequeDataBoa(ByVal pDataProcessamento As Long) As String

    Dim sStr    As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM ChequeDataBoa"
    sStr = sStr & " WHERE DataProcessamento < " & Trim(Str(pDataProcessamento))
    
    LimpaChequeDataBoa = sStr

End Function
Public Function LimpaAlteracaoData(ByVal pDataProcessamento As Long) As String

    Dim sStr    As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM AlteracaoData"
    sStr = sStr & " WHERE DataNova < " & Trim(Str(pDataProcessamento))
    
    LimpaAlteracaoData = sStr

End Function
Public Function LimpaChequesBaixados(ByVal pDataProcessamento As Long) As String

    Dim sStr    As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM ChequesBaixados"
    sStr = sStr & " WHERE DataDeposito < " & Trim(Str(pDataProcessamento))
    
    LimpaChequesBaixados = sStr

End Function
Public Function LimpaParametro(ByVal pDataProcessamento As Long) As String

    Dim sStr        As String
        
    sStr = "DELETE"
    sStr = sStr & " FROM Parametro"
    sStr = sStr & " WHERE DataProcessamento < " & Trim(Str(pDataProcessamento))
    
    LimpaParametro = sStr

End Function

Public Function RemoveGrupoUsuario(IdUsuario As Integer) As String

    Dim sStr  As String
    
    sStr = "DELETE"
    sStr = sStr & " From GRUPOUSUARIO where IdUsuario = " & IdUsuario
    RemoveGrupoUsuario = sStr

End Function
Public Function RemoveUsuario(IdUsuario As Integer) As String

    Dim sStr  As String
    
    sStr = "DELETE from USUARIO WHERE IdUsuario = " & IdUsuario
    RemoveUsuario = sStr

End Function
Public Function RemoveCheque(ByVal pDataProcessamento As Long, _
                             ByVal pIdBordero As Long, _
                             ByVal pIdcheque As Long) As String
    
    Dim sStr As String
    
           sStr = " DELETE FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento < " & pDataProcessamento
    sStr = sStr & " AND Idbordero =" & pIdBordero
    sStr = sStr & " AND Idcheque  =" & pIdcheque

    RemoveCheque = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Remove Cheques da Tabela Cheques pela Data de Deposito'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveCheques(ByVal pDataProcessamento As Long, _
                              ByVal pIdBordero As Long, _
                              ByVal pDataDeposito As Long) As String

    Dim sStr As String
    
           sStr = "DELETE"
    sStr = sStr & " FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento < " & pDataProcessamento
    sStr = sStr & " AND Idbordero =" & pIdBordero
    sStr = sStr & " AND DataDeposito = " & pDataDeposito

    RemoveCheques = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Remove registros de inconsistência por campo chave (Data e Hora Chave)          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveInconsistencia(ByVal strDataChave As String, ByVal strHoraChave As String) As String

    Dim sStr As String, strDataPosterior As String, dtDataPosterior As Date

   'Calcula 30 dias anteriores da DataChave para eliminar registros "perdidos" na base
    dtDataPosterior = Right(strDataChave, 2) & "/" & Mid(strDataChave, 5, 2) & "/" & Left(strDataChave, 4)
    dtDataPosterior = dtDataPosterior - 30
    strDataPosterior = Format(dtDataPosterior, "yyyymmdd")
    
    sStr = "Delete from Inconsistencia "
    sStr = sStr & "where (DataChave = '" & strDataChave & "' "
    sStr = sStr & "and HoraChave = '" & strHoraChave & "') "
    sStr = sStr & "or VAL(DataChave) <= " & Val(strDataPosterior)
    
    RemoveInconsistencia = sStr

End Function
Public Function RemoveComplementoExportacao(ByVal sReferencia As String) As String

     Dim sStr As String

     sStr = "Delete from ComplementoExportacao where Referencia = '" & sReferencia & "'"

     RemoveComplementoExportacao = sStr

End Function
Public Function RemoveCamposExportacao(ByVal sReferencia As String) As String

     Dim sStr As String

     sStr = "Delete from CamposExportacao where Referencia = '" & sReferencia & "'"

     RemoveCamposExportacao = sStr

End Function

