VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryMDI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Banco         As String
Dim m_Servidor      As String
Dim m_Usuario       As String
Dim m_Senha         As String
Dim m_Provedor      As String

Const Modulo = "Classe DbMDI"

Public DBMDI        As New ADODB.Connection
Public Property Get Banco() As String
    Banco = m_Banco
End Property
Public Property Let Banco(New_Banco As String)
    m_Banco = New_Banco
End Property
Public Property Get Servidor() As String
    Servidor = m_Servidor
End Property
Public Property Let Servidor(New_Servidor As String)
    m_Servidor = New_Servidor
End Property
Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property
Public Property Let Usuario(New_Usuario As String)
    m_Usuario = New_Usuario
End Property
Public Property Get Senha() As String
    Senha = m_Senha
End Property
Public Property Let Senha(New_Senha As String)
    m_Senha = New_Senha
End Property
Public Property Get Provedor() As String
    Provedor = m_Provedor
End Property
Public Property Let Provedor(New_Provedor As String)
    m_Provedor = New_Provedor
End Property
Function insConax(confscdprod As String, confscdfebr As String, _
                  confscdsegu As String, confsaxprod As String, _
                  confstparre As String, confsstrepa As String, _
                  confsnoprod As String, confsstbarr As String) As Integer
                  
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into CONAX "
    sStr = sStr & "(confscdprod,"
    sStr = sStr & " confscdfebr,"
    sStr = sStr & " confscdsegu,"
    sStr = sStr & " confsaxprod,"
    sStr = sStr & " confstparre,"
    sStr = sStr & " confsstrepa,"
    sStr = sStr & " confsnoprod,"
    sStr = sStr & " confsstbarr) "
    sStr = sStr & " VALUES ( '"
    sStr = sStr & confscdprod & "', '"
    sStr = sStr & confscdfebr & "', '"
    sStr = sStr & confscdsegu & "', '"
    sStr = sStr & confsaxprod & "', '"
    sStr = sStr & confstparre & "', '"
    sStr = sStr & confsstrepa & "', '"
    sStr = sStr & confsnoprod & "', '"
    sStr = sStr & confsstbarr & "')"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insConax
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "Query - InsConax"), Err.Description
    
End Function
Function insABGAG(ByVal abgfscdprod As Integer, _
                  ByVal abgfsabagen As String, _
                  ByVal abgfscdagen As Integer) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into ABGAG "
    sStr = sStr & "(abgfscdprod, "
    sStr = sStr & "abgfsabagen, "
    sStr = sStr & "abgfscdagen) "
    sStr = sStr & "VALUES ( "
    sStr = sStr & abgfscdprod & ", '"
    sStr = sStr & abgfsabagen & "', "
    sStr = sStr & abgfscdagen & ")"
    
    Set cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insABGAG
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "Query - InsABAG"), Err.Description
         
End Function
Function insAgenf(ByVal agefsnoagen As String, _
                  ByVal agefsestado As String, _
                  ByVal agefscdagen As String, _
                  ByVal agefsstmovi As String, _
                  ByVal agefsdtmvan As String, _
                  ByVal agefsdtmvat As String, _
                  ByVal agefsdtprox As String) As Integer
                  
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
           sStr = "Insert into AGENF "
    sStr = sStr & "(agefsnoagen, "
    sStr = sStr & "agefsestado, "
    sStr = sStr & "agefscdagen, "
    sStr = sStr & "agefsstmovi, "
    sStr = sStr & "agefsdtmvan, "
    sStr = sStr & "agefsdtmvat, "
    sStr = sStr & "agefsdtprox) "
    sStr = sStr & "VALUES ('"
    sStr = sStr & Replace(Trim(agefsnoagen), "'", "", 1, Len(agefsnoagen), vbTextCompare) & "', '"
    sStr = sStr & Left(agefsestado, 2) & "', "
    sStr = sStr & agefscdagen & ", "
    sStr = sStr & agefsstmovi & ", "
    sStr = sStr & agefsdtmvan & ", "
    sStr = sStr & agefsdtmvat & ", "
    sStr = sStr & agefsdtprox & ")"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insAgenf
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsAgenf"), Err.Description
          
End Function
Function insValidaGare(recfscodigo As String, recfscgrupo As String) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
    sStr = "Insert into ValidaGARE "
    sStr = sStr & "(Codigo, Grupo)"
    sStr = sStr & " VALUES ('"
    sStr = sStr & recfscodigo & "', '"
    sStr = sStr & recfscgrupo & "')"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insValidaGare
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsValidaGare"), Err.Description
        
End Function
Function insTfsBanco(banfscdbanc As String, banfsnobanc As String) As Integer
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
            
    sStr = "Insert into TFSBANCO "
    sStr = sStr & "(Codigo, "
    sStr = sStr & "Descricao) "
    sStr = sStr & "VALUES ('"
    sStr = sStr & banfscdbanc & "' , '"
    sStr = sStr & banfsnobanc & "')"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insTfsBanco
    
    Exit Function
    
TrataErro:
        If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsTfsBanco"), Err.Description

End Function
Function insTfsCred(crefsnubinc As String, crefscdbunc As String, _
                    crefsdescri As String, crefscdsaqu As String, _
                    crefscdagen As String, crefsnuccor As String, _
                    crefsbandei As String) As String
                    
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into TFSCCRED "
    sStr = sStr & "(crefsnubinc,"
    sStr = sStr & " crefscdbunc,"
    sStr = sStr & " crefsdescri,"
    sStr = sStr & " crefscdsaqu,"
    sStr = sStr & " crefscdagen,"
    sStr = sStr & " crefsnuccor,"
    sStr = sStr & " crefsbandei)"
    sStr = sStr & " VALUES ('"
    sStr = sStr & crefsnubinc & "', '"
    sStr = sStr & crefscdbunc & "', '"
    sStr = sStr & Trim(crefsdescri) & "', '"
    sStr = sStr & crefscdsaqu & "', '"
    sStr = sStr & crefscdagen & "', '"
    sStr = sStr & crefsnuccor & "', '"
    sStr = sStr & crefsbandei & "')"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insTfsCred
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsTfsCre"), Err.Description
        
End Function
Function insValidaGps(pDetalheGps As String) As String

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into ValidaGPS "
    sStr = sStr & "(CodigoPagamento,"
    sStr = sStr & " TipoDocumento,"
    sStr = sStr & " DigitaValorInss,"
    sStr = sStr & " DigitaOutrasEntidades,"
    sStr = sStr & " VerificaAtraso,"
    sStr = sStr & " DataInicial,"
    sStr = sStr & " DataFinal,"
    sStr = sStr & " IndicMesCompetencia,"
    sStr = sStr & " LimiteInicialCompetencia,"
    sStr = sStr & " IndicRestricaoPagto,"
    sStr = sStr & " DataPagtoNormal,"
    sStr = sStr & " CondPagtoNormal,"
    sStr = sStr & " DataPagto13,"
    sStr = sStr & " CondPagto13)"
    sStr = sStr & " VALUES ("
    sStr = sStr & Mid(pDetalheGps, 1, 4) & ", "
    sStr = sStr & "'" & Mid(pDetalheGps, 5, 1) & "', "
    sStr = sStr & "'" & Mid(pDetalheGps, 6, 1) & "', "
    sStr = sStr & "'" & Mid(pDetalheGps, 7, 1) & "', "
    sStr = sStr & "'" & Mid(pDetalheGps, 8, 1) & "', "
    sStr = sStr & Mid(Mid(pDetalheGps, 9, 8), 5, 4) & _
    Mid(Mid(pDetalheGps, 9, 8), 3, 2) & Mid(Mid(pDetalheGps, 9, 8), 1, 2) & ", "
    sStr = sStr & Mid(Mid(pDetalheGps, 17, 8), 5, 4) & _
    Mid(Mid(pDetalheGps, 17, 8), 3, 2) & Mid(Mid(pDetalheGps, 17, 8), 1, 2) & ",'"
    sStr = sStr & Mid(pDetalheGps, 25, 1) & "', "
    sStr = sStr & Mid(pDetalheGps, 26, 4) & ",'"
    sStr = sStr & Mid(pDetalheGps, 30, 1) & "','"
    sStr = sStr & Mid(pDetalheGps, 31, 2) & "','"
    sStr = sStr & Mid(pDetalheGps, 33, 1) & "',"
    sStr = sStr & Mid(pDetalheGps, 34, 4) & ",'"
    sStr = sStr & Mid(pDetalheGps, 38, 1) & "')"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insValidaGps
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsValidaGPS"), Err.Description
    
End Function
Function insValidaDarfPreto(pDetalheDarf As String, _
                            pDarfExcDocto As String, _
                            pDarfData As Long) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
       
    sStr = "Insert into ValidaDarfPreto "
    sStr = sStr & "(CodigoReceita,"
    sStr = sStr & " TipoDocumento,"
    sStr = sStr & " Referencia,"
    sStr = sStr & " DataInicial,"
    sStr = sStr & " DataFinal,"
    sStr = sStr & " ValorLimite,"
    sStr = sStr & " AnoLimiteInicial,"
    sStr = sStr & " AnoLimiteFinal,"
    sStr = sStr & " ExcessaoDocto,"
    sStr = sStr & " DataLimiteApur)"
    sStr = sStr & " VALUES ("
    sStr = sStr & Mid(pDetalheDarf, 1, 4) & ", "
    sStr = sStr & "'" & Mid(pDetalheDarf, 5, 1) & "', "
    sStr = sStr & "'" & Mid(pDetalheDarf, 6, 1) & "', "
    sStr = sStr & Mid(pDetalheDarf, 7, 8) & ", "
    sStr = sStr & Mid(pDetalheDarf, 15, 8) & ", "
    sStr = sStr & Replace(Mid(pDetalheDarf, 23, 16) / 100, ",", ".", 1, 10) & ", "
    sStr = sStr & Mid(pDetalheDarf, 39, 4) & ", "
    sStr = sStr & Mid(pDetalheDarf, 43, 4) & ", "
    sStr = sStr & "'" & pDarfExcDocto & "', "
    sStr = sStr & Format(pDarfData, "0000") & Mid(pDetalheDarf, 49, 2) & _
    Mid(pDetalheDarf, 47, 2) & ")"
        
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insValidaDarfPreto
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsValidaDarfPreto"), Err.Description
        
End Function
Function insRegraMdi_GrupoGare(pCodigoGrupoReceita As String, _
                               pIndicadorCotaIPVA As String, _
                               pIndicadorVenctoNormal As String, _
                               pIndicadorInscEstadual As String, _
                               pIndicadorCampoDocto As String, _
                               pIndicadorInscAtiva As String, _
                               pIndicadorReferencia As String, _
                               pIndicadorNumParcelamento As String, _
                               pIndicadorValorReceita As String, _
                               pIndicadorValorJuro As String, _
                               pIndicadorValorMulta As String, _
                               pIndicadorAcresFinanceiro As String, _
                               pIndicadorHonoAdvogado As String) As Integer
                               
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into REGRAMDI_GRUPOGARE (CodigoGrupo, "
    sStr = sStr & "Cota_IPVA, "
    sStr = sStr & "Vecto_Normal,  "
    sStr = sStr & "Inscricao_Estadual,  "
    sStr = sStr & "Campo_Docto, "
    sStr = sStr & "Inscricao_Ativa, "
    sStr = sStr & "Referencia, "
    sStr = sStr & "Numero_Parcelamento, "
    sStr = sStr & "Valor_Receita, "
    sStr = sStr & "Valor_Juros, "
    sStr = sStr & "Valor_Multa, "
    sStr = sStr & "Valor_Acrescimo_Financ, "
    sStr = sStr & "Valor_Honor_Advoc) "
    sStr = sStr & " VALUES ('"
    sStr = sStr & pCodigoGrupoReceita & "', "
    sStr = sStr & pIndicadorCotaIPVA & ", "
    sStr = sStr & pIndicadorVenctoNormal & ", "
    sStr = sStr & pIndicadorInscEstadual & ", "
    sStr = sStr & pIndicadorCampoDocto & ", "
    sStr = sStr & pIndicadorInscAtiva & ", "
    sStr = sStr & pIndicadorReferencia & ", "
    sStr = sStr & pIndicadorNumParcelamento & ", "
    sStr = sStr & pIndicadorValorReceita & ", "
    sStr = sStr & pIndicadorValorJuro & ", "
    sStr = sStr & pIndicadorValorMulta & ", "
    sStr = sStr & pIndicadorAcresFinanceiro & ", "
    sStr = sStr & pIndicadorHonoAdvogado & ")"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insRegraMdi_GrupoGare
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "insRegraMdi_GrupoGare"), Err.Description
    
End Function
Function insRegraMdi_Gare(pCodigoPagamento As String, _
                          pCodigoGrupo As String, _
                          pData_Inicio_Vigencia As String, _
                          pData_Final_Vigencia As String, _
                          pIndicador_Excessao As String, _
                          pIndicador_Arrecadacao As String, _
                          pTipo_Servico As String, _
                          pIndicador_Autenticacao As String, _
                          pIndicador_Servico_Autenticacao As String, _
                          pNumero_Vias_Comprovante As String, _
                          pValor As String) As Integer
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into REGRAMDI_GARE (CodigoPagamento, "
    sStr = sStr & "CodigoGrupo , "
    sStr = sStr & "Data_Inicio_Vigencia , "
    sStr = sStr & "Data_Final_Vigencia , "
    sStr = sStr & "Indicador_Excessao , "
    sStr = sStr & "Indicador_Arrecadacao , "
    sStr = sStr & "Tipo_Servico , "
    sStr = sStr & "Indicador_Autenticacao , "
    sStr = sStr & "Indicador_Servico_Autenticacao , "
    sStr = sStr & "Numero_Vias_Comprovante , "
    sStr = sStr & "Valor) "
    sStr = sStr & "VALUES ("
    sStr = sStr & pCodigoPagamento & ", '"
    sStr = sStr & pCodigoGrupo & "', "
    sStr = sStr & pData_Inicio_Vigencia & ", "
    sStr = sStr & pData_Final_Vigencia & ", '"
    sStr = sStr & pIndicador_Excessao & "', "
    sStr = sStr & pIndicador_Arrecadacao & ", "
    sStr = sStr & pTipo_Servico & ", '"
    sStr = sStr & pIndicador_Autenticacao & "', "
    sStr = sStr & pIndicador_Servico_Autenticacao & ", "
    sStr = sStr & pNumero_Vias_Comprovante & ", "
    sStr = sStr & pValor & ")"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Retorna Linhas afetadas
    cmd.Execute insRegraMdi_Gare
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "insRegra_Gare"), Err.Description
    
End Function

Function getDataSvr() As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
        
    sStr = "SELECT GETDATE() AS DataServMDI"
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getDataSvr = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "GetDataSvr"), Err.Description
    
End Function
Function getUsuario(ByVal pUsuario As String) As ADODB.Recordset

On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
        
    sStr = "Exec GetUsuario '" & pUsuario & "'"
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getUsuario = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "GetUsuário"), Err.Description

End Function
Function insLog(pDataprocessamento As String, _
                pIdCapa As Long, _
                pIdDocto As Long, _
                pUsuario As String, _
                pAcao) As Integer
                
On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command
        
    sStr = ""
    sStr = sStr & "Exec InsereLog "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & pUsuario & ", "
    sStr = sStr & pAcao
    
    Set cmd.ActiveConnection = DBMDI
    
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    insLog = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsereLog"), Err.Description
    
End Function
Function getCargaTabelas(ByVal pDataprocessamento As String) As ADODB.Recordset

On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
        
    sStr = "Exec GetCargaTabelas " & CLng(pDataprocessamento)
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getCargaTabelas = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "getCargaTabelas"), Err.Description

End Function
Function insParametro(pDataprocessamento As String) As Integer
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
   
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "CriarParametro"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
    
        .Execute
    End With
    
    insParametro = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "MDI - insParametro"), Err.Description
       
End Function
Function getParametro(ByVal pDataprocessamento As String) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
        
    sStr = "Exec LerParametro " & CLng(pDataprocessamento)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getParametro = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "GetParametro"), Err.Description
        
End Function
Function updParametroAlcada(ByVal VlLimInf As String, _
                            ByVal ValorAlcada_Env As String, _
                            ByVal ValorAlcada_Mal As String, _
                            ByVal PrazoVencimento_Env As String, _
                            ByVal PrazoVencimento_Mal As String, _
                            ByVal pDataprocessamento As String) As Integer
                            
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
           sStr = "UPDATE parametro set "
    sStr = sStr & "ValorInferior = " & formataValor(VlLimInf) & ", "
    sStr = sStr & "ValorAlcada_Env = " & ValorAlcada_Env & ", "
    sStr = sStr & "ValorAlcada_Mal = " & ValorAlcada_Mal & ", "
    sStr = sStr & "PrazoVencimento_Env = " & PrazoVencimento_Env & ", "
    sStr = sStr & "PrazoVencimento_Mal = " & PrazoVencimento_Mal
    sStr = sStr & " WHERE dataprocessamento = " & pDataprocessamento
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    
   'Executa e retorno "Rows Affecteds"
    cmd.Execute updParametroAlcada
    
    Exit Function
        
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "updParametroAlcada"), Err.Description

End Function
Function updCargaTabelas(ByVal pDataprocessamento As String) As Integer

On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command
    
    sStr = "Exec AtualizaCargaTabelas "
    sStr = sStr & CLng(pDataprocessamento) & ", 'S'"
        
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updCargaTabelas = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "updCargaTabelas"), Err.Description
    
End Function
Function updDoctoTransmitido(ByVal pDataprocessamento As String, _
                             ByVal pIdCapa As Long, _
                             ByVal pIdDocto As Long, _
                             ByVal pNSU As String, _
                             ByVal pCaixa As Long, _
                             ByVal pCortado As String) As Integer
                             
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
    sStr = "Exec RB_UpdDocumentoTransmitido "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", '"
    sStr = sStr & pNSU & "', "
    sStr = sStr & pCaixa & ", '"
    sStr = sStr & pCortado & "'"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updDoctoTransmitido = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdDocumentoTransmitido"), Err.Description
              
End Function
Function updNsuDocto(ByVal pDataprocessamento As String, _
                     ByVal pIdCapa As Long, _
                     ByVal pIdDocto As Long, _
                     ByVal pNSU1 As String, _
                     ByVal pCaixa As Long) As Integer
                     
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
    sStr = "exec RB_UpdNSUDocumento "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", '"
    sStr = sStr & pNSU1 & "', "
    sStr = sStr & pCaixa
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updNsuDocto = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdNSUDocumento"), Err.Description
   
End Function
Function getContraPartida(ByVal pEvento As Single) As ADODB.Recordset

On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
        
    sStr = "Exec RB_GetContraPartida "
    sStr = sStr & pEvento

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getContraPartida = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetContraPartida"), Err.Description
        
End Function
Function updAlteraTipoDocto(ByVal pDataprocessamento As String, _
                            ByVal pIdDocto As Long, _
                            ByVal pTipoDocto As Long) As Integer
                            
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
    sStr = "Exec AlteraTipoDocto "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto & ","
    sStr = sStr & pTipoDocto
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updAlteraTipoDocto = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "MDI - AlteraTipoDocto"), Err.Description
  
End Function
Function updCaixaDocto(ByVal pDataprocessamento As String, _
                       ByVal pCaixa As Long) As Integer
                  
On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command
    
    sStr = "Exec RB_UpdCaixaDocto "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pCaixa
        
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updCaixaDocto = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:

    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdCaixaDocto"), Err.Description

End Function
Function updCaixaCapa(ByVal pDataprocessamento As String, _
                      ByVal pCaixa As Long) As Integer
                  
On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command
    
    sStr = "Exec RB_UpdCaixaCapa "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pCaixa
        
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updCaixaCapa = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:

    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdCaixaDocto"), Err.Description

End Function
Function updDoctoRejeitado(ByVal pDataprocessamento As String, _
                           ByVal pIdCapa As Long, _
                           ByVal pIdDocto As Long, _
                           ByVal pOcorrencia As Integer, _
                           ByVal pTransacao As Integer, _
                           ByVal pStatus As String) As Integer
On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command
    
    sStr = "Exec RB_UpdDocumentoRejeitado "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & Format(pOcorrencia, "000") & ", "
    sStr = sStr & Format(pTransacao, "00") & ", '"
    sStr = sStr & pStatus & "'"
        
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updDoctoRejeitado = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdDocumentoRejeitado"), Err.Description

End Function
Function updDoctosRejeitados(ByVal pDataprocessamento As String, _
                             ByVal pIdCapa As Long, _
                             ByVal pIdDocto As Long, _
                             ByVal pVinculo As Long, _
                             ByVal pOcorrencia As Long, _
                             ByVal pTransacao As Single, _
                             Optional ByRef pQuantidade As Integer) As Integer
                    
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "RB_UpdDocumentosRejeitados"
    
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
    cmd.Parameters.Append cmd.CreateParameter("IdCapa", adDouble, adParamInput, 8, CDbl(pIdCapa))
    cmd.Parameters.Append cmd.CreateParameter("IdDocto", adDouble, adParamInput, 8, CDbl(pIdDocto))
    cmd.Parameters.Append cmd.CreateParameter("Vinculo", adDouble, adParamInput, 8, CDbl(pVinculo))
    cmd.Parameters.Append cmd.CreateParameter("Ocorrencia", adDouble, adParamInput, 8, CDbl(pOcorrencia))
    cmd.Parameters.Append cmd.CreateParameter("Transacao", adDouble, adParamInput, 8, CDbl(pTransacao))
    cmd.Parameters.Append cmd.CreateParameter("Quantidade", adDouble, adParamOutput, 8, CDbl(pQuantidade))
    
    cmd.Execute
    
    updDoctosRejeitados = cmd.Parameters("Retorno").Value
    pQuantidade = cmd.Parameters("Quantidade").Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdDocumentosRejeitados"), Err.Description

End Function
Function getTipoDocto(ByVal pTipoDocto As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "Exec GetTipoDocto " & pTipoDocto
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getTipoDocto = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "GetTipoDocto"), Err.Description
    
End Function
Function updStatusEstorno(ByVal pDataprocessamento As String, _
                          ByVal pIdCapa As Long, _
                          ByVal pIdDocto As Long, _
                          ByVal pStatus As String) As Integer

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "RB_UpdStatusEstorno"
    
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
    cmd.Parameters.Append cmd.CreateParameter("IdCapa", adDouble, adParamInput, 8, CDbl(pIdCapa))
    cmd.Parameters.Append cmd.CreateParameter("IdDocto", adDouble, adParamInput, 8, CDbl(pIdDocto))
    cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, pStatus)
    
    cmd.Execute
    
    updStatusEstorno = cmd.Parameters("Retorno").Value
    
    Exit Function
    
TrataErro:
    
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdStatusEstorno"), Err.Description
    
End Function
Function getConsultaCobrancasUBB(ByVal pDataprocessamento As String, _
                                 ByVal pIdCapa As Long, _
                                 ByVal pVinculo As Long) As ADODB.Recordset
                                 
    Dim sStr    As String
    Dim rs      As ADODB.Recordset
    
    sStr = "Exec RB_ConsultaCobrancaUBB "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getConsultaCobrancasUBB = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "getConsultaCobrancasUBB"), Err.Description
    
End Function
Function getConsultaCobrancaUBB(ByVal pProcedure As String, _
                                ByVal pDataprocessamento As String, _
                                ByVal pIdDocto As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
    
    sStr = "Exec " & pProcedure & " "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getConsultaCobrancaUBB = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "getConsultaCobrancaUBB"), Err.Description
    
End Function
Function updAtualizaOcorrenciaBHVR(ByVal pDataprocessamento As String, _
                                   ByVal pIdDocto As Long, _
                                   ByVal descBHVR As String) As Integer
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Exec RB_UPDOcorrenciaBHVR "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & "'" & descBHVR & "'"
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updAtualizaOcorrenciaBHVR = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:

    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_AtualizaOcorrenciaBHVR"), Err.Description

End Function
Function insAjusteDaCobranca(ByVal pDataprocessamento As String, _
                             ByVal pIdCapa As String, _
                             ByVal pTipoDocto As String, _
                             ByVal pAgencia As String, _
                             ByVal pConta As String, _
                             ByVal pValor As String, _
                             ByVal pVinculo As String, _
                             ByVal pAjNovaOrdem) As Integer
On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command

    sStr = "Exec InsereAjuste "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pTipoDocto & ", "
    sStr = sStr & pAgencia & ", "
    sStr = sStr & pConta & ", "
    sStr = sStr & formataValor(pValor) & ", "
    sStr = sStr & pVinculo & ", "
    sStr = sStr & pAjNovaOrdem

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    insAjusteDaCobranca = cmd.Parameters(0).Value

    Exit Function

TrataErro:

    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "insAjusteDaCobranca"), Err.Description
 
End Function
Function getUltimaOrdemCaptura(ByVal pDataprocessamento As String, ByVal pIdCapa As Long) As ADODB.Recordset
    
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset
    
    sStr = "Exec GetUltimaOrdemCaptura "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & CLng(pDataprocessamento)
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getUltimaOrdemCaptura = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "GetUltimaOrdemCaptura"), Err.Description

    
End Function
Function getCapaPendente(ByVal pDataprocessamento As String, _
                         ByVal pCaixa As Long) As ADODB.Recordset
                         
    On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetCapaPendente"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("Caixa", adChar, adParamInput, 12, pCaixa)
        
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
    End With
    
    Set getCapaPendente = rs
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_getCapaPendente"), Err.Description

End Function
Function getCapaRecepcao(ByVal pDataprocessamento As String, Optional ByRef pRetorno As Integer) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    Static Process
    
    With cmd
    
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_getCapaRecepcao"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
                
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
        
    End With
    
    If cmd.Parameters(0).Value = 0 And rs.State = adStateClosed Then
        If Process = Empty Then rs.Requery
        Process = True
    End If

    If rs.EOF Or cmd.Parameters(0).Value <> 0 Or rs.State = adStateClosed Then
        pRetorno = 1
    Else
        pRetorno = 0
        Set getCapaRecepcao = rs
    End If
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_getCapaRecepcao"), Err.Description
                           
End Function
Function updStatusCapa(ByVal pDataprocessamento As String, _
                       ByVal pIdCapa As Long, _
                       ByVal pStatus As String) As Integer
                       
On Error GoTo TrataErro

    Dim cmd As New ADODB.Command
    Dim sStr As String

    sStr = "exec AtualizaStatusCapa "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", '"
    sStr = sStr & pStatus & "'"
      
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updStatusCapa = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdStatusCapa"), Err.Description
End Function
Function updStatusDocto(ByVal pDataprocessamento As String, _
                        ByVal pIdDocto As Long, _
                        ByVal pStatus As String) As Integer
                       
On Error GoTo TrataErro

    Dim cmd As New ADODB.Command
    Dim sStr As String

    sStr = "exec AtualizaStatusDocumento "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto & ", '"
    sStr = sStr & pStatus & "'"
      
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updStatusDocto = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "AtualizaStatusDocumento"), Err.Description

End Function
Function getDocumentoEstorno(ByVal pDataprocessamento As String, _
                             ByVal pEstacao As String) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    
    With cmd
    
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetDocumentoEstorno"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("Caixa", adChar, adParamInput, 4, pEstacao)
                
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
        
    End With
    
    Set getDocumentoEstorno = rs
   
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetDocumentoEstorno"), Err.Description
                           
End Function
Function delTabelas(ByVal pTabela As String) As Integer

On Error GoTo TrataErro

    Dim cmd As New ADODB.Command
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = "Truncate Table " & pTabela
    cmd.Execute delTabelas
    
    Exit Function
    
TrataErro:

    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
      
    Err.Raise Err.Number, SetErrSource(Modulo, "Limpeza de Tabelas Auxiliares"), Err.Description
    
End Function
Function getCaixaNSU(ByVal pDataprocessamento As String, _
                     ByVal pEstacao As String, _
                     ByVal pCaixaInicial As Integer, _
                     ByRef pCaixa As Integer, _
                     ByRef pNSU As Integer) As Integer

On Error GoTo TrataErro
    
    Dim sStr    As String
    Dim cmd     As New ADODB.Command
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetCaixaNSU"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("Estacao", adChar, adParamInput, 5, pEstacao)
        .Parameters.Append cmd.CreateParameter("CaixaInicial", adInteger, adParamInput, 4, pCaixaInicial)
        .Parameters.Append cmd.CreateParameter("Caixa", adInteger, adParamOutput, 4, pCaixa)
        .Parameters.Append cmd.CreateParameter("NSU", adInteger, adParamOutput, 6, pNSU)
        .Execute
    End With
        
    If cmd.Parameters("Retorno").Value = 0 Then
        getCaixaNSU = 0
        pCaixa = cmd.Parameters("Caixa").Value
        pNSU = cmd.Parameters("NSU").Value
    End If
   
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(Modulo, "getCaixaNSU"), Err.Description

End Function
Function AtualizarCaixaMDI(ByVal pDataprocessamento As String, _
                           ByVal pEstacao As String, _
                           ByVal pCaixa As Long, _
                           ByVal pAbreFecha As String) As Integer
                        
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_AtualizarCaixaMDI"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("Estacao", adChar, adParamInput, 12, pEstacao)
        .Parameters.Append cmd.CreateParameter("Caixa", adDouble, adParamInput, 8, pCaixa)
        .Parameters.Append cmd.CreateParameter("AbreFecha", adChar, adParamInput, 1, pAbreFecha)
    
        .Execute
    End With

    AtualizarCaixaMDI = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_FecharCaixaMDI"), Err.Description
End Function
Function getChecaCapa(ByVal pDataprocessamento As String, _
                      ByVal pIdCapa As Long, _
                      ByRef pRetorno As Integer) As ADODB.Recordset

On Error GoTo TrataErro
    
    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    Static Process
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetChecaCapa"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("IdCapa", adDouble, adParamInput, 8, pIdCapa)
        
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
    End With
    
    If cmd.Parameters(0).Value = 0 And rs.State = adStateClosed Then
        If Process = Empty Then rs.Requery
        Process = True
    End If
    
    If cmd.Parameters(0).Value <> 0 Or rs.State = adStateClosed Then
        pRetorno = 1
    Else
        pRetorno = 0
    End If
    
    Set getChecaCapa = rs
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
 
    Err.Raise Err.Number, SetErrSource(Modulo, "getChecaCapa"), Err.Description

End Function
Function UpdCapaStatusCaixaControle(ByVal pDataprocessamento As String, _
                                    ByVal pIdCapa As Long, _
                                    ByVal pStatus As String, _
                                    ByVal pCaixa As Long) As Integer
                            
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String

    sStr = "exec RB_UpdCapaStatusCaixaControle "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", '"
    sStr = sStr & pStatus & "', "
    sStr = sStr & pCaixa
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    UpdCapaStatusCaixaControle = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdCapaStatusCaixaControle"), Err.Description
        
End Function
Function getAgenciaContaAjustes(ByVal pDataprocessamento As String, _
                                ByVal pIdCapa As Long, _
                                ByRef pRetorno As Integer) As ADODB.Recordset
    
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    Static Process
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetAgenciaContaDosAjustes"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("IdCapa", adDouble, adParamInput, 8, pIdCapa)
        
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
    End With
    
    If cmd.Parameters(0).Value = 0 And rs.State = adStateClosed Then
        If Process = Empty Then rs.Requery
        Process = True
    End If
    
    If cmd.Parameters(0).Value <> 0 Or rs.State = adStateClosed Then
        pRetorno = 1
    Else
        pRetorno = 0
        Set getAgenciaContaAjustes = rs
    End If
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetAgenciaContaDosAjustes"), Err.Description
    
End Function
Function updDoctoPendente(ByVal pDataprocessamento As String, _
                          ByVal pIdCapa As Long) As Integer
    
On Error GoTo TrataErro

    Dim sStr    As String
    Dim cmd     As New ADODB.Command
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_updDocumentoPendente"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("IdCapa", adDouble, adParamInput, 8, pIdCapa)
    
        .Execute
    End With
    
    updDoctoPendente = cmd.Parameters(0).Value

    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_updDocumentoPendente"), Err.Description

End Function
Function getDocumentos(ByVal pDataprocessamento As String, _
                       ByVal pIdCapa As Long, _
                       Optional ByRef pRetorno As Integer) As ADODB.Recordset

On Error GoTo TrataErro


    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    Static Process
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetDocumentos"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("IdCapa", adDouble, adParamInput, 8, pIdCapa)
        
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
    End With
    
    If cmd.Parameters(0).Value = 0 And rs.State = adStateClosed Then
        If Process = Empty Then rs.Requery
        Process = True
    End If
    
    If cmd.Parameters(0).Value <> 0 Or rs.State = adStateClosed Then
        pRetorno = 1
    Else
        pRetorno = 0
        Set getDocumentos = rs
    End If
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetDocumentos"), Err.Description
      
End Function
Function getDocumento(ByVal pDataprocessamento As String, _
                      ByVal pIdDocto As Long, _
                      ByVal pTipoDocto As Integer, _
                      ByVal pCaixa As Integer, _
                      Optional ByRef pRetorno As Integer) As ADODB.Recordset

On Error GoTo TrataErro


    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    Static Process

    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_GetDocumento"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("IdDocto", adDouble, adParamInput, 8, pIdDocto)
        .Parameters.Append cmd.CreateParameter("TipoDocto", adInteger, adParamInput, 2, pTipoDocto)
        .Parameters.Append cmd.CreateParameter("Caixa", adInteger, adParamInput, 4, pCaixa)
        
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
    End With
    
    If InStr(1, "05_06_08_09_20_21_22_23_24_25_26_42_43", Format(pTipoDocto, "00"), vbTextCompare) <> 0 And cmd.Parameters(0).Value = 0 Then
        pRetorno = 0
    Else
        If cmd.Parameters(0).Value = 0 And rs.State = adStateClosed Then
            If Process = Empty Then rs.Requery
            Process = True
        End If
        
        If cmd.Parameters(0).Value <> 0 Or rs.State = adStateClosed Then
            pRetorno = 1
        Else
            pRetorno = 0
            Set getDocumento = rs
        End If
    End If
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetDocumento"), Err.Description
      
End Function
Function getCapaTransmitir(ByVal pDataprocessamento As String, _
                           ByVal pCaixa As Long, Optional ByRef pRetorno As Integer) As ADODB.Recordset
    On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
    Dim cmd     As New ADODB.Command
    Static Process
    
    With cmd
        .ActiveConnection = DBMDI
        .CommandType = adCmdStoredProc
        .CommandText = "RB_getCapaTransmitir"
    
        .Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
        .Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
        .Parameters.Append cmd.CreateParameter("Caixa", adChar, adParamInput, 12, pCaixa)
        
        rs.CursorLocation = adUseClient
        
        Set rs = .Execute
    End With
    
    If cmd.Parameters(0).Value = 0 And rs.State = adStateClosed Then
        If Process = Empty Then rs.Requery
        Process = True
    End If

    If rs.EOF Or cmd.Parameters(0).Value <> 0 Or rs.State = adStateClosed Then
        pRetorno = 1
    Else
        pRetorno = 0
        Set getCapaTransmitir = rs
    End If
                
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetCapaTransmitir"), Err.Description

End Function
Function getAgenciaContadosDepositos(ByVal pDataprocessamento As String, _
                                     ByVal pIdCapa As Long, _
                                     ByRef pRetorno As Integer) As ADODB.Recordset

On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec RB_GetAgenciaContaDosDepositos "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF Or rs.State = adStateClosed Then
        pRetorno = 1
    Else
        pRetorno = 0
        Set getAgenciaContadosDepositos = rs
    End If

    Exit Function
    
TrataErro:

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetAgenciaContaDosDepositos"), Err.Description
     
End Function
Function updCancelarRetornoTransacao(ByVal pDataprocessamento As String, _
                                     ByVal pIdCapa As Long, _
                                     ByVal pIdDocto As Long, _
                                     ByVal pCaixa As Long) As Integer
                            
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    

    sStr = "exec RB_UpdCancelarRetornoTransacao "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & pCaixa

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updCancelarRetornoTransacao = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdCancelarRetornoTransacao"), Err.Description
            
End Function
Function updStatusChqOCT(ByVal pDataprocessamento As String, _
                         ByVal pIdCapa As Long, _
                         ByVal pVinculo As Long) As Integer
                       
On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String

    sStr = "exec RB_UpdStatusChqOCT "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo
      
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updStatusChqOCT = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "UpdStatusChqOCT"), Err.Description
End Function
Function getVerificaBinCartao(ByVal pCartao As String) As ADODB.Recordset

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec VerificaBinCartao '" & Mid(Trim(pCartao), 1, 6) & "'"
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getVerificaBinCartao = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "getVerificaBinCartao"), Err.Description

End Function
Function updCapaRecepcionada(ByVal pDataprocessamento As String, _
                             ByVal pIdCapa As Long, _
                             ByVal pStatus As String) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String

    sStr = "exec RB_UpdCapaRecepcionada "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ",'"
    sStr = sStr & pStatus & "'"

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updCapaRecepcionada = cmd.Parameters(0).Value
    
    Exit Function

TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_updCapaRecepcionada"), Err.Description
        
End Function
Function getTipoPagto(ByVal pDataprocessamento As String, _
                      ByVal pIdCapa As Long, _
                      ByVal pVinculo As Long) As ADODB.Recordset
                      
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec GetTipoPagto "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getTipoPagto = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "getTipoPagto"), Err.Description
    
End Function
Function getBancoOrigemDoAjuste(ByVal pDataprocessamento As String, _
                                ByVal pIdCapa As Long, _
                                ByVal pVinculo As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec GetBancoOrigemDoAjuste "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getBancoOrigemDoAjuste = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "GetBancoOrigemDoAjuste"), Err.Description

End Function
Function getChequesDeposito(ByVal pDataprocessamento As String, _
                            ByVal pIdCapa As Long, _
                            ByVal pVinculo As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec RB_GetChequesDeposito "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getChequesDeposito = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "GetChequesDeposito"), Err.Description
    
End Function
'Function getCMC7CapaOCT(ByVal pDataprocessamento As String, _
'                        ByVal pIdCapa As Long, _
'                        ByVal pVinculo As Long) As ADODB.Recordset
'On Error GoTo TrataErro
'
'    Dim sStr    As String
'    Dim rs      As ADODB.Recordset
'
'    sStr = "exec RB_GetCMC7CapaOCT "
'    sStr = sStr & CLng(pDataprocessamento) & ", "
'    sStr = sStr & pIdCapa & ", "
'    sStr = sStr & pVinculo
'
'    Set rs = New ADODB.Recordset
'
'    rs.CursorLocation = adUseClient
'    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText
'
'    Set getCMC7CapaOCT = rs
'
'    Exit Function
'
'TrataErro:
'    If Not rs Is Nothing Then
'        Set rs = Nothing
'    End If
'
'    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetCMC7CapaOCT"), Err.Description
'
'End Function
Function getOcorrenciasDeposito(ByVal pDataprocessamento As String, _
                                ByVal pIdCapa As Long, _
                                ByVal pVinculo As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec RB_GetOcorrenciasDeposito "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getOcorrenciasDeposito = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetOcorrenciasDeposito"), Err.Description

End Function
Function updOcorrenciaDocumento(ByVal pDataprocessamento As String, _
                                ByVal pIdDocto As Long, _
                                ByVal pCodOcorrencia As Single) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String

    sStr = sStr & "Exec RB_UpdOcorrenciaDocumento "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & pCodOcorrencia

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updOcorrenciaDocumento = cmd.Parameters(0).Value
    
    Exit Function

TrataErro:
        
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdOcorrenciaDocumento"), Err.Description
        
End Function
Function updHoraUltimaTransmissao(ByVal pDataprocessamento As String, _
                                  ByVal pCaixa As Long) As Integer

On Error GoTo TrataErro
    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "exec RB_UpdHoraUltimaTransmissao "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pCaixa

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updHoraUltimaTransmissao = cmd.Parameters(0).Value
    
    Exit Function

TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdHoraUltimaTransmissao"), Err.Description
End Function
Function insDiferencaCaixa(pDataprocessamento As String, _
                           pIdCapa As String, _
                           pCaixa As String, _
                           pEstacao As String, _
                           pValorDif As String) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
    
    sStr = "Insert into DiferencaCaixa "
    sStr = sStr & "(DataProcessamento,"
    sStr = sStr & " IdCapa,"
    sStr = sStr & " Caixa,"
    sStr = sStr & " IdEstacao,"
    sStr = sStr & " ValorDiferenca,"
    sStr = sStr & " Data_Hora) "
    sStr = sStr & " VALUES ("
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pCaixa & ", '"
    sStr = sStr & pEstacao & "', "
    sStr = sStr & pValorDif & ", "
    sStr = sStr & "GetDate()" & ")"
    
    Set cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Execute insDiferencaCaixa
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "Query - insDiferencaCaixa"), Err.Description
         
End Function
Function updChequeSaldo(ByVal pDataprocessamento As String, _
                        ByVal pIdDocto As Long, _
                        ByVal pMsgRetorno As String) As Integer

    On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
                                           
    sStr = "RB_UpdChequeSaldo "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & formataValor(Val(Mid(pMsgRetorno, 114, 1) & Mid(pMsgRetorno, 115, 16)), True) & ", "
    sStr = sStr & formataValor(Mid(pMsgRetorno, 131, 16), True) & ", "
    sStr = sStr & formataValor(Mid(pMsgRetorno, 147, 16), True)

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updChequeSaldo = cmd.Parameters(0).Value
    
    Exit Function

TrataErro:
    Screen.MousePointer = 0
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdChequeSaldo"), Err.Description
        
End Function
Function updADCCSaldo(pDataprocessamento As String, _
                      pIdDocto As Long, _
                      pMsgRetorno) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String

    sStr = "RB_UpdADCCSaldo "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & formataValor(Val(Mid(pMsgRetorno, 114, 1) & Mid(pMsgRetorno, 115, 16)), True) & ", "
    sStr = sStr & formataValor(Mid(pMsgRetorno, 131, 16), True) & ", "
    sStr = sStr & formataValor(Mid(pMsgRetorno, 147, 16), True)

    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updADCCSaldo = cmd.Parameters(0).Value
    
    Exit Function

TrataErro:
   
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdADCCSaldo"), Err.Description
End Function
Function updGareAD(pDataprocessamento As String, pIdDocto As Long, pAD As String) As Integer

On Error GoTo TrataErro

    Dim cmd     As New ADODB.Command
    Dim sStr    As String

    sStr = "RB_UpdGareAutenticaoDigital "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdDocto & ", '"
    sStr = sStr & pAD & "' "
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updGareAD = cmd.Parameters(0).Value
    
    Exit Function

TrataErro:
   
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_updGareAD"), Err.Description
End Function
Function updChequeDepositoTransmitido(ByVal pDataprocessamento As String, _
                                      ByVal pIdCapa As Long, _
                                      ByVal pIdDocto As Long, _
                                      ByVal pNSU As String, _
                                      ByVal pCaixa As Long) As Integer

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
    sStr = "Exec RB_UpdChequeDepositoTransmitido "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", '"
    sStr = sStr & pNSU & "', "
    sStr = sStr & pCaixa
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    updChequeDepositoTransmitido = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_UpdChequeDepositoTransmitido"), Err.Description
                 
End Function
Public Function insLogErro(ByVal pDataprocessamento As String, _
                           ByVal pEstacao As String, _
                           ByVal pErrNumber As Long, _
                           ByVal pDescricao As String) As Integer
                           
On Error GoTo TrataErro:

    Dim cmd     As New ADODB.Command
    Dim sStr    As String
        
    sStr = "exec InsereLogErro "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & "'" & pEstacao & "', "
    sStr = sStr & "'Robo " & pEstacao & "', "
    sStr = sStr & Format(pErrNumber, "000") & ", "
    sStr = sStr & "'" & pDescricao & "'"
           
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    insLogErro = cmd.Parameters(0).Value
    
    Exit Function
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Execute
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsLogErro"), Err.Description

End Function
Function insMensagem(ByVal pDataprocessamento As String, _
                     ByVal pIdMsg As Long, _
                     ByVal pIdCapa As Long, _
                     ByVal pIdDocto As Long, _
                     ByVal pCaixa As Long) As Integer
                     
                            
    Dim sStr    As String
    Dim cmd     As New ADODB.Command
        
    sStr = "Exec RB_InsMensagem "
    sStr = sStr & pDataprocessamento & ", "
    sStr = sStr & pIdMsg & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pIdDocto & ", "
    sStr = sStr & pCaixa
    
    Set cmd.ActiveConnection = DBMDI
    
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    insMensagem = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "InsereMensagem"), Err.Description
      
End Function
Function insControleCapa(ByVal pDataprocessamento As String, _
                         ByVal pIdCapa As Long, _
                         ByVal pModulo As Integer) As Integer

    Dim cmd         As New ADODB.Command
    Dim pComentario As String
    Dim sStr        As String
    
    pComentario = "'Robo envia Capa p/ CSP'"
        
    sStr = "Exec InsereControleCapa "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pComentario & ", "
    sStr = sStr & pModulo
        
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = sStr
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Execute
    
    insControleCapa = cmd.Parameters(0).Value
    
    Exit Function
    
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "AtualizaControleCapa"), Err.Description

End Function
Function getUltimoNSU(ByVal pDataprocessamento As String, ByVal pCaixa As Long) As Integer
                    
On Error GoTo TrataErro

    Dim cmd As New ADODB.Command

    cmd.ActiveConnection = DBMDI
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "RB_GetUltimoNSU"
    
    cmd.Parameters.Append cmd.CreateParameter("Retorno", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("Data", adDouble, adParamInput, 8, CDbl(pDataprocessamento))
    cmd.Parameters.Append cmd.CreateParameter("Caixa", adDouble, adParamInput, 8, CDbl(pCaixa))
    cmd.Parameters.Append cmd.CreateParameter("NSU", adDouble, adParamOutput, 8, getUltimoNSU)
    
    cmd.Execute
    
    If cmd.Parameters("Retorno").Value = 0 Then
        getUltimoNSU = cmd.Parameters("NSU").Value
        Exit Function
    End If
        
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetUltimoNSU"), Err.Description

End Function
Function getZero() As ADODB.Recordset

On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "Select 0 from StatusCapa"
    
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getZero = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "GetZero"), Err.Description
    
End Function
Public Sub SetConnection()

On Error GoTo TrataErro:

    With DBMDI
        .ConnectionTimeout = 240
        .Provider = Me.Provedor
        .Properties("Data Source").Value = Servidor
        .Properties("Initial Catalog").Value = Banco
        
       'Decision code for login authorization type: WinNT or SQL Server.
        If Me.Usuario = "NT" Then
            .Properties("Integrated Security").Value = "SSPI"
        Else
            .Properties("User ID").Value = Me.Usuario
            .Properties("Password").Value = Me.Senha
        End If
        
       .CursorLocation = adUseClient
       .Open
        
    End With
    
    Exit Sub
    
TrataErro:
    If Err.Number = adErrObjectOpen Then
        Exit Sub
    End If
    
    If Not DBMDI Is Nothing Then
        Set DBMDI = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "Abertura de Conexão com MDI"), Err.Description
    
End Sub
Private Sub Class_Terminate()
    DBMDI.Close
End Sub
Function ExecuteSQL(sStr As String) As Recordset
On Error GoTo TrataErro

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set ExecuteSQL = rs
        
    Exit Function
    
TrataErro:
    Err.Raise Err.Number, SetErrSource(Modulo, "Função: ExecuteSQL"), Err.Description
    
End Function
Function ExecuteUPD(ByVal pStr As String) As Integer
                            
    Dim cmd As New ADODB.Command
    
    cmd.ActiveConnection = DBMDI
    cmd.CommandText = pStr
    
   'Executa e retorno "Rows Affecteds"
    cmd.Execute ExecuteUPD
    
    Exit Function
        
TrataErro:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "ExecuteUPD"), Err.Description

End Function
Function getComplOcorrencia(ByVal pDataprocessamento As String, _
                            ByVal pIdDocto As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As New ADODB.Recordset
        
    sStr = "exec RB_GetComplOcorrencia "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdDocto
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getComplOcorrencia = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
  
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetComplOcorrencia"), Err.Description
                           
End Function
Function getCMC7CapaOCT(ByVal pDataprocessamento As String, _
                        ByVal pIdCapa As Long, _
                        ByVal pVinculo As Long) As ADODB.Recordset
On Error GoTo TrataErro

    Dim sStr    As String
    Dim rs      As ADODB.Recordset

    sStr = "exec RB_GetCMC7CapaOCT "
    sStr = sStr & CLng(pDataprocessamento) & ", "
    sStr = sStr & pIdCapa & ", "
    sStr = sStr & pVinculo

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sStr, DBMDI, adOpenStatic, adLockReadOnly, adCmdText

    Set getCMC7CapaOCT = rs
    
    Exit Function
    
TrataErro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    Err.Raise Err.Number, SetErrSource(Modulo, "RB_GetCMC7CapaOCT"), Err.Description

End Function


