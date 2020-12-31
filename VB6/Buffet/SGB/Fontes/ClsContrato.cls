VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsContrato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function RegistrarChequeCompensado(ByVal ID_CNT As Integer, ByVal ID_PAR As Integer, ByVal Db As ADODB.Connection) As Integer

    On Error GoTo RegistrarChequeCompensado_Erro
    Dim sSql As String
    Dim Rec As New ADODB.Recordset

    sSql = "UPDATE PARCELA_CONTRATO SET COMPENSADO = 'S' "
    sSql = sSql & " WHERE ID_CNT = " & ID_CNT
    sSql = sSql & " AND ID_PAR = " & ID_PAR
    
    Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly

    Exit Function

RegistrarChequeCompensado_Erro:
    MsgBox "Erro ao registrar a compensação do Cheque", vbExclamation, "SGB"
End Function
Public Function Incluir(ID_CNT As String, ByVal Db As ADODB.Connection) As Integer

    Dim sSql As String
    Dim Rec As New ADODB.Recordset
    Dim x As Long
    Dim CodContrato As Integer

    On Error GoTo Incluir_Erro

    If Val(ID_CNT) = 0 Then
        'Recuperar o último código do contrato cadastrado
        sSql = "select max(id_cnt) as CONTRATO from CONTRATOS"

        Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly

        If Not Rec.EOF Then
            If IsNull(Rec("CONTRATO").Value) Then
                CodContrato = 1
            Else
                CodContrato = Rec("CONTRATO").Value
                CodContrato = CodContrato + 1
            End If
            Incluir = True
        Else
            Incluir = False
        End If
        Set Rec = Nothing
    Else
        CodContrato = Val(ID_CNT)
    End If

    sSql = "insert into CONTRATOS (ID_CNT,COD_CLI,NOME_ANIV,HR_INI,HR_FIM,DATA_FESTA,IDADE_ANIV,"
    sSql = sSql & "QTDE_CONV,DATA_CNT,VALOR_TOTAL,DATA_ANIV,CUSTO_ADIC,"
    sSql = sSql & "DSC_BOLO,DSC_DECOR,OBS_BEBIDA,NOM_PAIS,OBS) VALUES ("
    'ID CONTRATO
    sSql = sSql & CodContrato & ","
    'COD CLIENTE
    sSql = sSql & Contrato.TxtCodCliente.Text & ", '"
    'NOME ANIVER
    sSql = sSql & Contrato.TxtAniversariante.Text & "' , '"
    'HR INI
    sSql = sSql & Contrato.CboHoraInicio.ItemData(Contrato.CboHoraInicio.ListIndex) & "' , '"
    'HR FIM
    sSql = sSql & Contrato.CboHoraFim.ItemData(Contrato.CboHoraFim.ListIndex) & "','"
    'DATA FESTA
    sSql = sSql & Format(Contrato.TxtDiaFesta.Text, "dd/mm/yyyy") & "','"
    'IDADE ANIV
    sSql = sSql & Contrato.TxtIdade.Text & "',"
    'QTDE_CONV
    sSql = sSql & Contrato.TxtQtdeConvidados.Text & ",'"
    'DATA_CNT
    sSql = sSql & Contrato.TxtDataContrato.Text & "',"
    'VALOR_TOTAL
    sSql = sSql & Int(Contrato.TxtCustoTotal.Text) & ",'"
    'DATA_ANIV
    sSql = sSql & Format(Contrato.TxtDataAniversario.Text, "mm/dd/yyyy") & "',"
    'CUSTO_ADIC
    sSql = sSql & Int(Contrato.TxtAdicionalPessoa.Text) & ","
    'COND_PAGTO
    'sSql = sSql & Contrato.CboFormaPagto.ItemData(Contrato.CboFormaPagto.ListIndex) & ","
    'ID_BOLO
    'sSql = sSql & Contrato.CboBolo.ItemData(Contrato.CboBolo.ListIndex) & ",'"
    'DSC_BOLO
    sSql = sSql & "'" & Contrato.TxtBolo.Text & "','"
    'ID_DECOR
    'sSql = sSql & Contrato.CboDecor.ItemData(Contrato.CboDecor.ListIndex) & ",'"
    'DSC_DECOR
    sSql = sSql & Contrato.TxtDecoracao.Text & "',"
    'ID_DOCE
    'sSql = sSql & Contrato.CboDoce.ItemData(Contrato.CboDoce.ListIndex) & ","
    'ID_SALGADO
    'sSql = sSql & Contrato.CboSalgado.ItemData(Contrato.CboSalgado.ListIndex) & ",'"
    'OBS_BEBIDA
    sSql = sSql & "'" & Contrato.TxtBebida.Text & "','"
    'NOM_PAIS
    sSql = sSql & Contrato.TxtPais.Text & "','"
    'OBS
    sSql = sSql & Contrato.TxtOBS.Text & "')"

    Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly

    Contrato.LblNumeroContrato.Caption = Format(CodContrato, "00000")
    Contrato.LblNumeroContrato.Visible = True
    DoEvents

    Exit Function

Incluir_Erro:
    MsgBox Error
    MsgBox "Ocorreu um erro e o contrato não pode ser gravado. Feche a tela de CONTRATOS e reinicie a operação.", vbOKOnly, "SGB"
    Exit Function
End Function
Public Function Excluir(ID_CNT As String, ByVal Db As ADODB.Connection) As Integer

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo Excluir_Erro

    sSql = "DELETE FROM CONTRATOS WHERE ID_CNT = " & ID_CNT

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    Excluir = True

    Exit Function

Excluir_Erro:
    MsgBox "Erro ao Excluir os dados do Contrato", vbExclamation, "SGB"
    Excluir = False
End Function

