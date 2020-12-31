<TITLE>debug.asp</TITLE>
<body bgcolor="#FFFFFF">
<HTML>
<% 
Response.Write("<P>VARIAVEIS DO FORMULARIO:<br>")
Response.Write("-------------------------------<br>")

For Each Key in Request.Form
	Response.Write( Key & " = " & Request.Form(Key) & "<br>")
Next
 
Response.Write("<P>VARIAVEIS QUERY STRING:<br>")
Response.Write("------------------------------<br>")
For Each Key in Request.QueryString
	Response.Write( Key & " = " & Request.QueryString(Key) & "<br>")
Next

Response.Write("<P>VARIAVEIS TIPO COOKIE:<br>")
Response.Write("-----------------------------<br>")

For Each Key in Request.Cookies
	Response.Write( Key & " = " & Request.Cookies(Key) & "<br>")
Next

Response.Write("<P>VARIAVEIS DE SERVIDOR:<br>")
Response.Write("-----------------------------<br>")

For Each Key in Request.ServerVariables
	Response.Write( Key & " = " & Request.ServerVariables(Key) & "<br>")
Next
%>
</BODY>
</HTML>
