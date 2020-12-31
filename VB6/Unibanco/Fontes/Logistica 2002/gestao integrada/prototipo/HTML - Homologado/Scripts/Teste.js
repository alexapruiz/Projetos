<script language="JScript">
Array.prototype.IsArray = "true";
String.prototype.IsArray = "false";

function Inicializa()
{
  var Teste;
  Teste = Comunica("", "");
}


function Comunica(xmldoc, strOpcao)
{    
  var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  xmlhttp.Open("POST", "http://server01/sistema/p.asp?Opcao=" + strOpcao, false);
  xmlhttp.Send(xmldoc);
  return(xmlhttp.responseText);
}

function MontaArray(objDoc)
{
  var arResult, arVariant, intLinhas, intColunas, intCol, intRow, i;
  
  //arResult = new Array(objDoc.documentElement.childNodes.length - 1);
  arResult = new Array();
  for (i = 0; i < objDoc.documentElement.childNodes.length; i++)
  {
    if (objDoc.documentElement.childNodes[i].childNodes[0].nodeName == "Variant")
    {
      intLinhas = objDoc.documentElement.childNodes[i].childNodes[0].childNodes.length;
      intColunas = objDoc.documentElement.childNodes[i].childNodes[0].childNodes[0].childNodes.length;
      //arVariant = new Array(intColunas - 1, intLinhas - 1);
      arVariant = new Array();
      for (intCol = 0; intCol < intColunas; intCol++)
      {
        arVariant[intCol] = new Array();
        for (intRow = 0; intRow < intLinhas; intRow++)
        {
          arVariant[intCol][intRow] = objDoc.documentElement.childNodes[i].childNodes[0].childNodes[intRow].childNodes[intCol].text;
        }
      }
      arResult[i] = arVariant;
    }    
    else
    {
      arResult[i] = objDoc.documentElement.childNodes[i].childNodes[0].text;
    }
  }
return arResult;
}

function MontaXML(vntArray)
{
  var intRow, intCol, intPos;
  var strCol, strLinha, strVariant, strInstrucao;
  var NodePrincipal, NodeParametro, NodeTipo, NodeLinha, NodeColuna;
  
  var objDoc = new ActiveXObject("MSXML2.DOMDocument");

  strInstrucao = objDoc.createProcessingInstruction("xml", "version='1.0'");
  strInstrucao = objDoc.appendChild(strInstrucao);

  NodePrincipal = objDoc.createNode(1,"root", "");  

  for (intPos = 0; intPos <= vntArray.length - 1; intPos++)
  {
    NodeParametro = objDoc.createNode(1,"Parametro", "");

    if (vntArray[intPos].IsArray == "false")
    {
      NodeTipo = objDoc.createNode(1, "String", "");      
      NodeTipo.text = vntArray[intPos];
    }
    else
    {
      NodeTipo = objDoc.createNode(1, "Variant", "");
      for (intRow = 0; intRow <= vntArray[intPos][0].length - 1; intRow++)
      {
        strCol = "";
        NodeLinha = objDoc.createNode(1, "Linha", "");
        for (intCol = 0; intCol <= vntArray[intPos].length - 1; intCol++)
        {
          NodeColuna = objDoc.createNode(1, "Coluna", "");
          NodeColuna.text = vntArray[intPos][intCol][intRow];
          NodeColuna = NodeLinha.insertBefore(NodeColuna, null);
        }
        NodeLinha = NodeTipo.insertBefore(NodeLinha, null);
      }
    }
    NodeTipo = NodeParametro.insertBefore(NodeTipo, null);
    NodeParametro = NodePrincipal.insertBefore(NodeParametro, null);
  }

  NodePrincipal = objDoc.insertBefore(NodePrincipal, null);
  return objDoc.xml;
}


</script>