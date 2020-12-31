function LTrim(String)
{
   var i = 0;
   var j = String.length - 1;

  if (String == null)
    return (false);

  for (i = 0; i < String.length; i++)
  {
    if (String.substr(i, 1) != ' ' && String.substr(i, 1) != '\t')
      break;
  }

  if (i <= j)
    return (String.substr(i, (j+1)-i));
  else
    return ('');
}

function RTrim(String)
{
  var i = 0;
  var j = String.length - 1;

  if (String == null)
    return (false);

  for(j = String.length - 1; j >= 0; j--)
  {
    if (String.substr(j, 1) != ' ' && String.substr(j, 1) != '\t')
    break;
  }

  if (i <= j)
    return (String.substr(i, (j+1)-i));
  else
    return ('');
}

function Trim(String)
{
  if (String == null)
    return (false);

  return RTrim(LTrim(String));
}

function Left(String, Length)
{
  if (String == null)
    return (false);

  return String.substr(0, Length);
}

function Right(String, Length)
{
  if (String == null)
    return (false);

  var dest = '';
  for (var i = (String.length - 1); i >= 0; i--)
    dest = dest + String.charAt(i);

  String = dest;
  String = String.substr(0, Length);
  dest = '';

  for (var i = (String.length - 1); i >= 0; i--)
    dest = dest + String.charAt(i);

  return dest;
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

function IsNumeric( Str )
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