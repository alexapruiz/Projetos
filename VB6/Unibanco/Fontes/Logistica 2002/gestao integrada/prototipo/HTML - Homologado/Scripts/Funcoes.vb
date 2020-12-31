<%
Function MontaXML(vntArray, Nivel)
  Dim intRow, intCol, intPos, vntTeste(0)
  Dim strCol, strLinha, strVariant, strInstrucao
  Dim NodePrincipal, NodeParametro, NodeTipo, NodeLinha, NodeColuna
  
  Dim objDoc
  Set objDoc = CreateObject("MSXML2.DOMDocument")

  If Nivel = 0 Then
    Set strInstrucao = objDoc.createProcessingInstruction("xml", "version=""1.0""")
    Set strInstrucao = objDoc.appendChild(strInstrucao)

    Set NodePrincipal = objDoc.createNode(1, "root", "")
  End If

  For intPos = LBound(vntArray) To UBound(vntArray)
    If Nivel = 0 Then Set NodeParametro = objDoc.createNode(1, "Parametro", "")
    If Not IsArray(vntArray(intPos)) Then
      Set NodeTipo = objDoc.createNode(1, "String", "")
      NodeTipo.Text = vntArray(intPos)
    Else
      Set NodeTipo = objDoc.createNode(1, "Variant", "")
      For intRow = LBound(vntArray(intPos), 2) To UBound(vntArray(intPos), 2)
        strCol = ""
        Set NodeLinha = objDoc.createNode(1, "Linha", "")
        For intCol = LBound(vntArray(intPos), 1) To UBound(vntArray(intPos), 1)
          Set NodeColuna = objDoc.createNode(1, "Coluna", "")
          If IsArray(vntArray(intPos)(intCol, intRow)) Then
            vntTeste(0) = vntArray(intPos)(intCol, intRow)
            NodeColuna.Text = MontaXML(vntTeste, 1)
          Else
            NodeColuna.Text = vntArray(intPos)(intCol, intRow)
          End If
          Set NodeColuna = NodeLinha.appendChild(NodeColuna)
        Next
        Set NodeLinha = NodeTipo.appendChild(NodeLinha)
      Next
    End If
    If Nivel = 0 Then
      Set NodeTipo = NodeParametro.appendChild(NodeTipo)
      Set NodeParametro = NodePrincipal.appendChild(NodeParametro)
    End If
  Next

  If Nivel = 0 Then
    Set NodePrincipal = objDoc.appendChild(NodePrincipal)
  Else
    Set NodeTipo = objDoc.appendChild(NodeTipo)
  End If

  MontaXML = Replace(Replace(objDoc.xml, "&gt;", ">"), "&lt;", "<")

End Function

Function MontaArray(objDoc, Nivel)
  Dim arResult, arVariant, intLinhas, intColunas, intCol, intRow, i, objDoc2
  
  ReDim arResult(objDoc.documentElement.childNodes.length - 1)
  For i = 0 To objDoc.documentElement.childNodes.length - 1
    If objDoc.documentElement.childNodes(i).childNodes(0).nodeName = "Variant" Then
      intLinhas = objDoc.documentElement.childNodes(i).childNodes(0).childNodes.length
      intColunas = objDoc.documentElement.childNodes(i).childNodes(0).childNodes(0).childNodes.length
      ReDim arVariant(intColunas - 1, intLinhas - 1)
      For intRow = 0 To intLinhas - 1
        For intCol = 0 To intColunas - 1
          If objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).childNodes.length <> 0 Then
            If objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).childNodes(0).nodeName = "Variant" Then
              Set objDoc2 = Server.CreateObject("MSXML2.DOMDocument")
              objDoc2.loadXML ("<root><parametro>" & objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).childNodes(0).xml & "</parametro></root>")
              arVariant(intCol, intRow) = MontaArray(objDoc2, 1)
            Else
              if objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).Text = "True" Then
                arVariant(intCol, intRow) = True
              else
                arVariant(intCol, intRow) = objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).Text
              end if
            End If
          Else
            if objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).Text = "True" Then
              arVariant(intCol, intRow) = True
            else
              arVariant(intCol, intRow) = objDoc.documentElement.childNodes(i).childNodes(0).childNodes(intRow).childNodes(intCol).Text
            end if
          End If
        Next
      Next
      If Nivel = 0 Then
        arResult(i) = arVariant
      Else
        arResult = arVariant
      End If
    Else
      arResult(i) = objDoc.documentElement.childNodes(i).childNodes(0).Text
    End If
  Next
  MontaArray = arResult

End Function

Function FormataNumero(Valor, Digitos)
  Dim Ret, i
  for i = 1 to Digitos - len(valor)
    Ret = "0" & Ret
  next
  FormataNumero = Ret & Valor
End Function 

Function FormataNum( Num, Dec )
  Dim iNum, i, lFrac, iNPs
 
  If not isnumeric(Num) or not isnumeric(Dec) Then
    FormataNum = null
    Exit Function
  End If
  iNum  = int(Num)
  If Dec <> 0 Then
    lFrac =  cdbl( cdbl(Num) - int(Num) )
    If lFrac <> 0 Then
      lFrac = lFrac & "000000000000000000000000000000000"
    Else
      lFrac = "0,00000000000000000000000000000000"
    End If
    lFrac = right( lFrac , len( lFrac )-1 )
    lFrac = left( lFrac , Dec + 1)
  Else
    lFrac = ""
  End If
  For i=1 To Len(iNum) / 3
    iNPs = "." & Right(iNum , 3) & iNPs
    iNum = Left(iNum , Len(iNum)-3 )
  next
  If len(iNum) = 0 Then
    iNPs = Right(iNPs,Len(iNPs)-1)
  Else
    iNPs = iNum & iNPs 
  End If
  FormataNum = iNPs & lFrac 
end function

Function Inicializa()
  Dim Teste
  Teste = Comunica("", "")
End Function


Function Comunica(xmldoc, strOpcao)
  Dim xmlhttp
  Set xmlhttp = Server.CreateObject("Msxml2.XMLHTTP")
  xmlhttp.Open "POST", "processa.asp?Opcao=" + strOpcao, false
  xmlhttp.Send xmldoc
  Comunica = xmlhttp.responseText
End Function

Function RegistrarErro(Mensagem)
   Dim ArqPronto, ArqAcesso, ArqLog, Caminho
   Const ForReading = 1, ForWriting = 2
   
   ArqPronto = false
   
   Caminho = Request.ServerVariables("PATH_TRANSLATED")
   do while Right(Caminho, 1) <> "\"
     Caminho = Left(Caminho, len(Caminho) - 1)
   loop
   
   Set ArqAcesso = Server.CreateObject("Scripting.FileSystemObject")
   
   On Error Resume Next
   
   Err.Clear
   
   Set ArqLog = ArqAcesso.OpenTextFile(Caminho & "\Log\Log.txt", 8 , True)
   
   if Err.number <> 0 then
     Err.Clear
     Set ArqLog = ArqAcesso.CreateTextFile(Caminho & "\Log\Log.txt", False)
     if Err.number = 0 then
       ArqPronto = true
     end if
   else
     ArqPronto = true
   end if
 
   if ArqPronto then
     ArqLog.WriteLine Now & " - " & Session("Login") & " - " & Request.ServerVariables("REMOTE_HOST") & " - " & Application("Id_AplSist") & "|" & Session("Id_Modulo") & "|" & Session("Id_SubModulo") & " - " & Mensagem
     ArqLog.close
   end if
   
   Set ArqLog = Nothing
   Set ArqAcesso = nothing
End Function

%>