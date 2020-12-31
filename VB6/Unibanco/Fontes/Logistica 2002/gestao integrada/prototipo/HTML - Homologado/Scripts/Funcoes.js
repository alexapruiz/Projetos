<script language="JScript">
Array.prototype.IsArray = "true";
String.prototype.IsArray = "false";

function Menu( Mostra )
{
  if( Mostra )
  { 
    /* Mostrar Menu e Diminuir Página... */
    document.all.DMenuOpen.style.visibility  = "hidden" ; 
    document.all.DMenuClose.style.visibility = "visible" ;
    document.all.BG_UBB.style.left           = -159 ;     
  }
  else
  {
    /* Esconder Menu e Aumentar Página... */
    document.all.DMenuOpen.style.visibility  = "visible" ; 
    document.all.DMenuClose.style.visibility = "hidden" ;
    document.all.BG_UBB.style.left           = 0 ;     
  }
  parent.MenuOn( Mostra );
}

function Inicializa()
{
  var Teste;
  Teste = Comunica("", "");
}


function Comunica(xmldoc, strOpcao)
{    
  var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  xmlhttp.Open("POST", "processa.asp?Opcao=" + strOpcao, false);
  xmlhttp.Send(xmldoc);
  return(xmlhttp.responseText);
}

function MontaArray(objDoc, Nivel)
{
  var arResult, arVariant, intLinhas, intColunas, intCol, intRow, i;
  
  arResult = new Array();
  for (i = 0; i < objDoc.documentElement.childNodes.length; i++)
  {
    if (objDoc.documentElement.childNodes[i].childNodes[0].nodeName == "Variant")
    {
      intLinhas = objDoc.documentElement.childNodes[i].childNodes[0].childNodes.length;
      intColunas = objDoc.documentElement.childNodes[i].childNodes[0].childNodes[0].childNodes.length;
      arVariant = new Array();
      for (intCol = 0; intCol < intColunas; intCol++)
      {
        arVariant[intCol] = new Array();
        for (intRow = 0; intRow < intLinhas; intRow++)
        {
          if (objDoc.documentElement.childNodes[i].childNodes[0].childNodes[intRow].childNodes[intCol].childNodes.length != 0)
          {
            if (objDoc.documentElement.childNodes[i].childNodes[0].childNodes[intRow].childNodes[intCol].childNodes[0].nodeName == "Variant")
            {
              var objDoc2 = new ActiveXObject("MSXML2.DOMDocument");
              objDoc2.loadXML ("<root><parametro>" + objDoc.documentElement.childNodes[i].childNodes[0].childNodes[intRow].childNodes[intCol].childNodes[0].xml + "</parametro></root>");
              arVariant[intCol][intRow] = MontaArray(objDoc2, 1);
            }
            else
            {
              arVariant[intCol][intRow] = objDoc.documentElement.childNodes[i].childNodes[0].childNodes[intRow].childNodes[intCol].text;
            }
          }
          else
          {
            arVariant[intCol][intRow] = objDoc.documentElement.childNodes[i].childNodes[0].childNodes[intRow].childNodes[intCol].text;
          }

        }
      }
      if (Nivel == null)
      {
        arResult[i] = arVariant;
      }
      else
      {
        arResult = arVariant;
      }
    }    
    else
    {
      arResult[i] = objDoc.documentElement.childNodes[i].childNodes[0].text;
    }
  }
return arResult;
}

function MontaXML(vntArray, Nivel)
{
  var intRow, intCol, intPos;
  var strCol, strLinha, strVariant, strInstrucao, vntTeste = new Array();
  var NodePrincipal, NodeParametro, NodeTipo, NodeLinha, NodeColuna;
  
  var objDoc = new ActiveXObject("MSXML2.DOMDocument");

  if (Nivel == null)
  {
    strInstrucao = objDoc.createProcessingInstruction("xml", "version='1.0'");
    strInstrucao = objDoc.appendChild(strInstrucao);

    NodePrincipal = objDoc.createNode(1,"root", "");  
  }


  for (intPos = 0; intPos <= vntArray.length - 1; intPos++)
  {
    if (Nivel == null)
    {
      NodeParametro = objDoc.createNode(1,"Parametro", "");
    }
    
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

          if (vntArray[intPos][intCol][intRow].IsArray == "true")
          {
            vntTeste[0] = vntArray[intPos][intCol][intRow];
            NodeColuna.text = MontaXML(vntTeste, 1);
          }
          else
    {
      NodeColuna.text = vntArray[intPos][intCol][intRow];
    }
          
          NodeColuna = NodeLinha.insertBefore(NodeColuna, null);
        }
        NodeLinha = NodeTipo.insertBefore(NodeLinha, null);
      }
    }
    if (Nivel == null)
    {
      NodeTipo = NodeParametro.insertBefore(NodeTipo, null);
      NodeParametro = NodePrincipal.insertBefore(NodeParametro, null);
    }
  }
  if (Nivel == null)
  {
    NodePrincipal = objDoc.insertBefore(NodePrincipal, null);
  }
  else
  {
    NodeTipo = objDoc.insertBefore(NodeTipo, null);
  }

  var cond1 = /&lt;/g;
  var cond2 = /&gt;/g;
  
  return objDoc.xml.replace(cond1, "<").replace(cond2, ">");
}

function IsNum(passedVal)
{
  if (passedVal == "")
  {
    return false;
  }
  for (i = 0; i < passedVal.length; i++)
  {
    if (passedVal.charAt(i) < "0")
    {
      return false;
    }
    if (passedVal.charAt(i) > "9")
        {
          return false;
    }
  }
  return true;
}


function turnOn(imageName)
{
  if (document.images)
  {
    document[imageName].src = eval(imageName + "on.src");
    document.body.style.cursor = "hand"
  }
}

function turnOff(imageName)
{
  if (document.images)
  {
    document[imageName].src = eval(imageName + "off.src");
    document.body.style.cursor = "default"
  }
}

function PreLoadImages()
{
  for (i = 0; i < document.images.length; i++)
  {
    eval(document.images(i).name + "on = new Image()");
    eval(document.images(i).name + "on.src = 'Imagens/" + document.images(i).name + "B.gif'");
    eval(document.images(i).name + "off = new Image()");
    eval(document.images(i).name + "off.src = 'Imagens/" + document.images(i).name + ".gif'");
  }
}

function Mid(String, Start, Length)
{
  if (String == null)
    return (false);

  if (Start > String.length)
    return '';

  if (Length == null || Length.length == 0)
    return (false);

  return String.substr((Start - 1), Length);
}


function IsNumeric(Str)
{
  var i;
  var char;
  
  for( i = 0; i < Str.length ; i++ )
  {
    char = Mid( Str, i, 1);
    if( char != "0" && char != "1"  && char != "2"  && char != "3"  && char != "4"  && 
        char != "5" && char != "6"  && char != "7"  && char != "8"  && char != "9"      )
      return false;
  }
  return true;
}

</script>